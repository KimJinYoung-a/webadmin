<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
''�Ե����� ��ǰ���� ��ȸ :: �Ǹſ��ε��� ��Ȯ�� �Ѿ���� �ʴµ�.., StockQty: ��ǰ�� �Ѿ����, ���տɼ���..
function LotteOneItemCheckStock(iitemid,byRef iErrStr)
    Dim ilottegoods_no
    Dim objXML,xmlDOM,strRst, iMessage
    Dim ProdCount, buf, AssignedRow, oneProdInfo, strParam
    Dim GoodNo,ItemNo,OptDesc,DispYn,SaleStatCd,StockQty, bufopt
    Dim strSql, actCnt


    On Error Resume Next
    LotteOneItemCheckStock = False

    ilottegoods_no =getTenItem2LotteGoodNo(iitemid)

    if (ilottegoods_no="") then
        iErrStr = "["&iitemid&"] �Ե� ���� �ڵ� ����."
        Exit function
	end if

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")

    strParam = "?subscriptionId=" & lotteAuthNo											'�Ե����� ������ȣ	(*)
    strParam = strParam & "&search_gubun=goods_no"
	strParam = strParam & "&search_text=" & ilottegoods_no		'�Ե����� ��ǰ��ȣ	(*)

''rw lotteAPIURL & "/openapi/searchStockList.lotte" & strParam

	objXML.Open "GET", lotteAPIURL & "/openapi/searchStockList.lotte"&strParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

    ''rw "objXML.Status="&objXML.Status
	If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��

		buf = BinaryToText(objXML.ResponseBody, "euc-kr")

	''CALL XMLFileSaveLotte(buf,"STQ",ilottegoods_no)  ''slow

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML replace(buf,"&","��")
'rw buf
		if Err<>0 then
		    iErrStr =  "�Ե����� ��� �м� �߿� ������ �߻� [" & iitemid & "]:"
            Set objXML = Nothing
            Set xmlDOM = Nothing
		    Exit function

			''Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			''dbget.Close: Response.End
		end if

        ProdCount   = Trim(xmlDOM.getElementsByTagName("ProdCount").item(0).text)   '' ��ǰ ����

       ''rw "ProdCount="&ProdCount
        IF (ProdCount="1") then         ''��ǰ�ΰ�� SKIP ==> ����
            if (getItemOptionCount(iitemid)<1) then ProdCount=""
        end if
    ''rw "aaaa"
        if (ProdCount<>"") then
            Set oneProdInfo = xmlDOM.getElementsByTagName("ProdInfo")

            strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption='')"
            strSql = strSql & " BEGIN"
            strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption=''"
            strSql = strSql & " END"
            dbget.Execute strSql

            ''2013/05/30 �߰�
            strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
            strSql = strSql & " BEGIN"
            strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6"
            strSql = strSql & " END"
            dbget.Execute strSql

            for each SubNodes in oneProdInfo
				GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
                ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' ��ǰ�ڵ� (���� 0,1,2,)
                OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
                DispYn	    = Trim(SubNodes.getElementsByTagName("DispYn").item(0).text)         ''N:���� Y:����
                SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''�Ǹ�����, �Ǹ�����, ǰ��'
                StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)

                OptDesc = replace(OptDesc,"��","&")
                if (SaleStatCd<>"�Ǹ�����") THEN
                    DispYn="N"
                END IF

                if (StockQty="null") then
                    StockQty="0"
                end if

                bufopt = OptDesc
                If InStr(bufopt,",")>0 then
                    if (splitValue(bufopt,",",0)<>"") then
                        OptDesc = splitValue(splitValue(bufopt,",",0),":",1)
                    end if

                    if (splitValue(bufopt,",",1)<>"") then
                        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",1),":",1)
                    end if

                    if (splitValue(bufopt,",",2)<>"") then
                        OptDesc = OptDesc+","+splitValue(splitValue(bufopt,",",2),":",1)
                    end if
                ELSE
                    OptDesc = splitValue(OptDesc,":",1)
                ENd IF
    ''rw "OptDesc="&OptDesc
				strSql = " update oP"
		        strSql = strSql & " set outmallOptName='"&html2DB(OptDesc)&"'"&VbCRLF
		        strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
		        strSql = strSql & " ,outMallSellyn='"&DispYn&"'"&VbCRLF
		        strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
		        strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
		        strSql = strSql & "     From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
		        strSql = strSql & " where itemid="&iitemid&VbCRLF
		        strSql = strSql & " and outmallOptCode='"&ItemNo&"'"&VbCRLF
		        strSql = strSql & " and mallid='lotteCom'"&VbCRLF
       ''rw strSql
		        dbget.Execute strSql, AssignedRow
		   ''rw "AssignedRow="&AssignedRow

		        if (AssignedRow<1) then
		            strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption='')"
		            strSql = strSql & " BEGIN"
		            strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and itemoption=''"
		            strSql = strSql & " END"
		            dbget.Execute strSql

		            strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
		            strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo)"
		            strSql = strSql & " values("&iitemid
		            strSql = strSql & " ,'"&ItemNo&"'" ''�ӽ÷� �Ե� �ڵ� ���� //2013/04/01
		            strSql = strSql & " ,'lotteCom'"
		            strSql = strSql & " ,'"&ItemNo&"'"
		            strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
		            strSql = strSql & " ,'"&DispYn&"'"
			        strSql = strSql & " ,'Y'"
			        strSql = strSql & " ,"&StockQty
			        strSql = strSql & ")"

			        dbget.Execute strSql, AssignedRow
			     ''rw "AssignedRow="&AssignedRow
			        ''�ɼ� �ڵ� ��Ī.
			        if (AssignedRow>0) then
			            strSql = " update oP"   &VbCRLF
			            strSql = strSql & " set itemoption=O.itemoption"&VbCRLF
			            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
			            strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
			            strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
			            strSql = strSql & " where oP.mallid='lotteCom'"&VbCRLF
			            strSql = strSql & " and o.itemid="&iitemid&VbCRLF
			            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
			            strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
			            strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
			            ''strSql = strSql & " and op.outmallOptName=o.optionname"&VbCRLF

			            dbget.Execute strSql, AssignedRow

			        end if
		        else
		            strSql = " update oP"   &VbCRLF
		            strSql = strSql & " set itemoption=O.itemoption"&VbCRLF
		            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
		            strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
		            strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
		            strSql = strSql & " where oP.mallid='lotteCom'"&VbCRLF
		            strSql = strSql & " and o.itemid="&iitemid&VbCRLF
		            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
		            strSql = strSql & " and op.outmallOptCode='"&ItemNo&"'"&VbCRLF
		            strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','')=Replace(Replace(o.optionname,' ',''),':','')"&VbCRLF
		       'rw strSql
		            dbget.Execute strSql, AssignedRow
		        end if

			    actCnt = actCnt+AssignedRow
		    next

			if (actCnt>0) then
			    strSql = " update R"   &VbCRLF
                strSql = strSql & " set regedOptCnt=isNULL(T.CNT,0)"   &VbCRLF
                strSql = strSql & " from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
                strSql = strSql & " 	Join ("   &VbCRLF
                strSql = strSql & " 		select R.itemid,count(*) as CNT from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
                strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
                strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
                strSql = strSql & " 			and Ro.mallid='lotteCom'"   &VbCRLF
                strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
                strSql = strSql & " 		group by R.itemid"   &VbCRLF
                strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF

                dbget.Execute strSql
			end if
		end if
		LotteOneItemCheckStock =true
		Set xmlDOM = Nothing
	else
	    iErrStr =  "�Ե����İ� ����߿� ������ �߻� [" & iitemid & "]:"
        Set objXML = Nothing
        Set xmlDOM = Nothing
	    Exit function

		''Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
		''dbget.Close: Response.End
	end if
    Set objXML = Nothing

    On Error Goto 0
end function

'����ī�װ� ����
Function LotteOneItemCateEdit(iitemid, strParam, byRef iErrStr)
	Dim objXML, xmlDOM, strRst, iMessage, strSql
	On Error Resume Next
	LotteOneItemCateEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsCategoryOpenApi.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		'rw lotteAPIURL & "/openapi/updateGoodsCategoryOpenApi.lotte?"&strParam
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				If Err <> 0 Then
				    iErrStr = "�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.LotteOneItemCateEdit"
					dbget.Close: Exit Function
				End If

				'��� �ڵ�
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// ���� ����
				If Err <> 0 Then
		            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
					dbget.Close: Exit Function
		        End If
			Set xmlDOM = Nothing
		Else
		    iErrStr ="�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''�Ե����� ��ǰ ���
function LotteOneItemReg(iitemid,strParam,byRef iErrStr,iSellCash,iLotteSellYn)
    Dim objXML,xmlDOM,strRst, iMessage
    Dim buf, LotteGoodNo, strSql, buf_item_list, pp, OptDesc, StockQty, AssignedRow

    On Error Resume Next
    LotteOneItemReg = False

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", lotteAPIURL & "/openapi/registApiGoodsInfo.lotte", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)
	If objXML.Status = "200" Then
''rw lotteAPIURL & "/openapi/registApiGoodsInfo.lotte?"&strParam
		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End
        buf = BinaryToText(objXML.ResponseBody, "euc-kr")
		''CALL XMLFileSaveLotte(buf,"REG",oLotteitem.FItemList(i).FItemID)

		''�����ڵ�� �����..(�ɼǸ� ��)

		LotteGoodNo = ""
		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML buf ''BinaryToText(objXML.ResponseBody, "euc-kr")

		LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

		'// ���� ����(��ǰ��ȣ�� �ݵ�� �����ؾ� ��)
		if LotteGoodNo="" then
		    iMessage = xmlDOM.getElementsByTagName("Message").item(0).text
		    iErrStr =  "��ǰ ����� ���� [" & iitemid & "]:"&iMessage
            Set objXML = Nothing
            Set xmlDOM = Nothing
		    Exit function
		end if

		if Err<>0 then
			iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if


		'��ǰ���翩�� Ȯ��
		strSql = "Select count(itemid) From db_item.dbo.tbl_lotte_regItem Where itemid='" & iitemid & "'"
		rsget.Open strSql,dbget,1

		if rsget(0)>0 then
			'// ���� -> ����
			strSql = "update R" & VbCRLF
			strSql = strSql & "	Set LotteLastUpdate=getdate() "  & VbCRLF
			strSql = strSql & "	, LotteTmpGoodNo='" & LotteGoodNo & "'"  & VbCRLF
			strSql = strSql & "	, LottePrice=" &iSellCash& VbCRLF
			strSql = strSql & "	, accFailCnt=0"& VbCRLF
			strSql = strSql & "	, lotteRegdate=isNULL(lotteRegdate,getdate())" ''�߰� 2013/02/26
			if (LotteGoodNo<>"") then
			    strSql = strSql & "	, lottestatCD='20'"& VbCRLF
			else
    			strSql = strSql & "	, lottestatCD='10'"& VbCRLF
    		end if
			strSql = strSql & "	From db_item.dbo.tbl_lotte_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid='" & iitemid & "'"

			dbget.Execute(strSql)
		else
			'// ���� -> �űԵ��
			strSql = "Insert into db_item.dbo.tbl_lotte_regItem "
			strSql = strSql & " (itemid, reguserid, lotteRegdate, LotteLastUpdate, LotteTmpGoodNo, LottePrice, LotteSellYn, LotteStatCd) values " & VbCRLF
			strSql = strSql & " ('" & iitemid & "'" & VbCRLF
			strSql = strSql & ", '" & session("ssBctId") & "'" &_
			strSql = strSql & ", getdate(), getdate()" & VbCRLF
			strSql = strSql & ", '" & LotteGoodNo & "'" & VbCRLF
			strSql = strSql & ", '" & iSellCash & "'" & VbCRLF
			strSql = strSql & ", '" & iLotteSellYn & "'" & VbCRLF
			if (LotteGoodNo<>"") then
			    strSql = strSql & ",'20'"
			else
			    strSql = strSql & ",'10'"
			end if
			strSql = strSql & ")"
			dbget.Execute(strSql)

			actCnt = actCnt+1
		end if

		rsget.Close

        IF (TRUE) THEN
            pp = 0
        ''�ɼ� ����Ʈ ����(20120807 �߰�)
            ''buf_item_list = xmlDOM.getElementsByTagName("item_list").item(0).text
            ''buf_item_list = xmlDOM.getElementsByTagName("Arguments").item("item_list").text
            IF xmlDOM.getElementsByTagName("Argument").item(38).getAttribute("name")="item_list" THEN
                buf_item_list = xmlDOM.getElementsByTagName("Argument").item(38).getAttribute("value")
            ELSE
                buf_item_list = xmlDOM.getElementsByTagName("Argument").item(39).getAttribute("value")
            END IF
            '''iitemid = oLotteitem.FItemList(i).Fitemid
            if (buf_item_list<>"") then
                buf_item_list = UniToHanbyChilkat(buf_item_list)

                rw "["&iitemid&"]=="&LotteGoodNo&"=="&buf_item_list

                buf_item_list = split(buf_item_list,":")
                for k=Lbound(buf_item_list) to Ubound(buf_item_list)
                    ''���߿ɼ��ΰ�� �� ���� �߸� �Ǿ���.
                    OptDesc = splitvalue(buf_item_list(k),",",0)
                    StockQty = splitvalue(buf_item_list(k),",",1)

                    strSql = " Insert into db_item.dbo.tbl_OutMall_regedoption"
		            strSql = strSql & " (itemid,itemoption,mallid,outmallOptCode,outmallOptName,outMallSellyn,outmalllimityn,outMallLimitNo)"
		            strSql = strSql & " values("&iitemid
		            strSql = strSql & " ,''"
		            strSql = strSql & " ,'lotteCom'"
		            strSql = strSql & " ,'"&pp&"'"
		            strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
		            strSql = strSql & " ,'Y'"
			        strSql = strSql & " ,'Y'"
			        strSql = strSql & " ,"&StockQty
			        strSql = strSql & ")"

			        dbget.Execute strSql, AssignedRow

			        ''�ɼ� �ڵ� ��Ī.
			        if (AssignedRow>0) then
			            strSql = " update oP"   &VbCRLF
			            strSql = strSql & " set itemoption=O.itemoption"&VbCRLF
			            strSql = strSql & " From db_item.dbo.tbl_OutMall_regedoption oP"&VbCRLF
			            strSql = strSql & "     Join db_item.dbo.tbl_item_option o"&VbCRLF
			            strSql = strSql & "     on oP.itemid=o.itemid"&VbCRLF
			            strSql = strSql & " where oP.mallid='lotteCom'"&VbCRLF
			            strSql = strSql & " and o.itemid="&iitemid&VbCRLF
			            strSql = strSql & " and oP.itemid="&iitemid&VbCRLF
			            strSql = strSql & " and op.outmallOptCode='"&pp&"'"&VbCRLF
			            strSql = strSql & " and op.outmallOptName=o.optionname"&VbCRLF

			            dbget.Execute strSql, AssignedRow

			        end if

			        pp = pp + 1
			    Next

			    strSql = " update R"   &VbCRLF
                strSql = strSql & " set regedOptCnt=isNULL(T.CNT,0)"   &VbCRLF
                strSql = strSql & " from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
                strSql = strSql & " 	Join ("   &VbCRLF
                strSql = strSql & " 		select R.itemid,count(*) as CNT from db_item.dbo.tbl_lotte_regItem R"   &VbCRLF
                strSql = strSql & " 			Join db_item.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
                strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
                strSql = strSql & " 			and Ro.mallid='lotteCom'"   &VbCRLF
                strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
                strSql = strSql & " 		group by R.itemid"   &VbCRLF
                strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF

                dbget.Execute strSql
            else
                rw "["&iitemid&"]=="&LotteGoodNo
	        end if
        END IF
		Set xmlDOM = Nothing

		LotteOneItemReg= true
	else
		iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-REG-002]"
	end if
	Set objXML = Nothing
    On Error Goto 0

end function

''�Ե����� ��ǰ���� ����
function LotteOneItemInfoEdit(iitemid,strParam,byRef iErrStr,isVer2)
    Dim objXML,xmlDOM,strRst, iMessage

    On Error Resume Next
    LotteOneItemInfoEdit = False

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    if (isVer2) then
        objXML.Open "POST", lotteAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''��ǰ����
    else
        objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''���û�ǰ����
    end if

    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send(strParam)

	If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
''response.end

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
		    iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
			''Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			''dbget.Close: Response.End
		end if

		'��� �ڵ�

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

        if (strRst<>"1") then
            ''rw "FitemDiv="&oLotteitem.FItemList(i).FitemDiv  ''�ֹ����� ��ǰ�ΰ�� ������ �ȵǴµ�. add_choc_tp_cd_20 �� ���� �ִ°��;;
            ''rw "iMessage="&iMessage
            iErrStr =  "��ǰ ������ ���� [" & iitemid & "]:"&iMessage
            Set objXML = Nothing
            Set xmlDOM = Nothing
            On Error Goto 0
		    Exit function
        end if

        if (strRst="1") then
			'// ���� ����
			if Err<>0 then
			    if (IsAutoScript) then
			        rw "["&iitemid & "]:"&iMessage
			    else
					iErrStr = "["&iitemid & "]:"&iMessage
				end if
				Set objXML = Nothing
    		    Set xmlDOM = Nothing
    		    Exit function
			else

			end if
        end if

		Set xmlDOM = Nothing
		LotteOneItemInfoEdit = True
	else
	    iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-EDIT-002]"
	end if
	Set objXML = Nothing

	On Error Goto 0
end function

function LotteOnItemPriceEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage

  On Error Resume Next
    LotteOnItemPriceEdit = False

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)

	If objXML.Status = "200" Then

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
			iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRCEDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'��� �ڵ�

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

		'// ���� ����
		if Err<>0 then
		    IF (Trim(iMessage)="002. ��������(�Ǵ� ���࿹��)") THEN
		        ''skip
		        iErrStr = "["&iitemid & "]:"&Trim(iMessage)&"\n"
		        Set objXML = Nothing
    		    Set xmlDOM = Nothing
    		    On Error Goto 0
    		    Exit function
		    ELSE
		        if (IsAutoScript) then
		            rw "["&iitemid & "]:"&iMessage
		        else
					iErrStr =  "["&iitemid & "]:"&iMessage
				end if
				Set objXML = Nothing
    		    Set xmlDOM = Nothing
    		    On Error Goto 0
    		    Exit function
			End IF
        else

        end if
		Set xmlDOM = Nothing
		LotteOnItemPriceEdit = True
	else
		iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-PRCEDIT-002]"
	end if
	Set objXML = Nothing
 On Error Goto 0
end function

function LotteOneItemSellStatEdit(iitemid,iLotteGoodNo,ichgSellYn,byRef iErrStr)
    Dim strParam
    Dim objXML, xmlDOM
    Dim strRst, strSql

    LotteOneItemSellStatEdit = False
	'//��ǰ��� �Ķ����
	strParam = "?subscriptionId=" & lotteAuthNo											'�Ե����� ������ȣ	(*)
	strParam = strParam & "&goods_no=" & iLotteGoodNo                       			'�Ե����� ��ǰ��ȣ	(*)
	if ichgSellYn="Y" then																'�Ǹſ���(10:�Ǹ�, 20:ǰ��, 30:�Ǹ�����)
		strParam = strParam & "&sale_stat_cd=10"
	elseif ichgSellYn="N" then
		strParam = strParam & "&sale_stat_cd=20"
	elseif ichgSellYn="X" then                           '''X ��� ������
		''''strParam = strParam & "&sale_stat_cd=30"      ''�Ǹ�����Ǹ� ��������.
		strParam = strParam & "&sale_stat_cd=20"
	end if

''rw lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
		    iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����"
			''Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			dbget.Close: Exit function
		end if

		'��� �ڵ�
		on Error resume next
		strRst = xmlDOM.getElementsByTagName("Result").item(0).text
        on Error Goto 0

		'// ���� ����
		if Err<>0 then
'           IF (xmlDOM.getElementsByTagName("Message").item(0).text="�Ǹ������� ��ǰ�� �ǸŻ��¸� ������ �� �����ϴ�.") THEN
'''         '// ��ǰ���� ���� 2012/03/06 ==> �ּ����� 2013/02/25
'''         strSql = "Update db_item.dbo.tbl_lotte_regItem " & VbCRLF
'''         strSql = strSql & " Set LotteSellYn='N'" & VbCRLF
'''         strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
'''         dbget.Execute(strSql)
'
'           Err=0
'       ELSE
            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
			''Response.Write "<script language=javascript>alert('Error: " & xmlDOM.getElementsByTagName("Message").item(0).text  & "');</script>"
			dbget.Close: Exit function
'''     END IF
		ELSE

			'// ��ǰ���� ����
			strSql = "Update db_item.dbo.tbl_lotte_regItem " & VbCRLF
			strSql = strSql & " Set LotteLastUpdate=getdate() " & VbCRLF
			strSql = strSql & " ,LotteSellYn='" & ichgSellYn & "'" & VbCRLF
			strSql = strSql & " Where itemid='" & iitemid & "'"
			dbget.Execute(strSql)
        end if


		Set xmlDOM = Nothing
	else
	    iErrStr ="�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����"
		''Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
		dbget.Close: Exit function
	end if
	Set objXML = Nothing

	LotteOneItemSellStatEdit = true
end function

function LotteOneItemNameEdit(iitemid,strParam,byRef iErrStr)
    'http://openapidev.lotte.com/openapi/updateGoodsNmOpenApi.lotte
    Dim objXML,xmlDOM,strRst,iMessage

  On Error Resume Next
    LotteOneItemNameEdit = False

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", lotteAPIURL & "/openapi/updateGoodsNmOpenApi.lotte", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)

	If objXML.Status = "200" Then

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
			iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'��� �ڵ�

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

		'// ���� ����
		if Err<>0 then
	        if (IsAutoScript) then
	            rw "["&iitemid & "]:"&iMessage
	        else
				iErrStr =  "["&iitemid & "]:"&iMessage
			end if
			Set objXML = Nothing
		    Set xmlDOM = Nothing
		    On Error Goto 0
		    Exit function
        else

        end if
		Set xmlDOM = Nothing
		LotteOneItemNameEdit = True
	else
		iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-002]"
	end if
	Set objXML = Nothing
 On Error Goto 0
end function

'ǰ������
function LotteOneItemPoomOkEdit(iitemid,strParam,byRef iErrStr)
	Dim objXML,xmlDOM,strRst,iMessage

	On Error Resume Next
	LotteOneItemPoomOkEdit = False
	
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsItemInfo.lotte", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err <> 0 Then
					iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PoomEDIT-001]"
					Set objXML = Nothing
					Set xmlDOM = Nothing
					Exit Function
				End If
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
				If strRst <> 1 Then
					iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				Else
					rw "["&iitemid & "]:ǰ������ ���� ����"
				End IF
				
				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
					Set xmlDOM = Nothing
					On Error Goto 0
					Exit Function
				End If
				Set xmlDOM = Nothing
				LotteOneItemPoomOkEdit = True
		Else
			iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-PoomEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
end function

''���û�ǰ ����ȸ
Function CheckLotteItemStatWithOption(iitemid,byRef iErrStr)
	Dim ilottegoods_no
	Dim objXML, xmlDOM
	Dim buf, oneProdInfo, strParam
	Dim ItemNo, InvQty, Opt1Nm, Opt1Tval, Opt2Nm, Opt2Tval, ItemSaleStatCd
	On Error Resume Next
	CheckLotteItemStatWithOption = False
	ilottegoods_no = getTenItem2LotteGoodNo(iitemid)
	If (ilottegoods_no = "") then
		iErrStr = "["&iitemid&"] �Ե� ���� �ڵ� ����."
		Exit Function
	End If
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strParam = "?subscriptionId=" & lotteAuthNo											'�Ե����� ������ȣ	(*)
		strParam = strParam & "&strGoodsNo=" & ilottegoods_no
		objXML.Open "GET", lotteAPIURL & "/openapi/searchGoodsViewListOpenApi.lotte"&strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

	If objXML.Status = "200" Then
		buf = BinaryToText(objXML.ResponseBody, "euc-kr")
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML replace(buf,"&","��")
''rw buf
		If Err <> 0 Then
			iErrStr =  "�Ե����� ��� �м� �߿� ������ �߻� [" & iitemid & "]:"
			Set objXML = Nothing
			Set xmlDOM = Nothing
			Exit function
		End If

			Set oneProdInfo = xmlDOM.getElementsByTagName("ItemInfo")
							'���տɼ����� 5������ �Ǿ��ִ� �� ������ 10x10�� ���߿ɼ��� 2���� �ɼ��ڵ�,�ɼǸ��� 2������ ������ ����
			For each SubNodes in oneProdInfo
				ItemNo			= Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)			'��ǰ�ɼǹ�ȣ
				InvQty			= Trim(SubNodes.getElementsByTagName("InvQty").item(0).text)			'�ɼ�������
				Opt1Nm 			= Trim(SubNodes.getElementsByTagName("Opt1Nm").item(0).text)			'�ɼ��ڵ�1
				Opt1Tval		= Trim(SubNodes.getElementsByTagName("Opt1Tval").item(0).text)			'�ɼǸ�1
				Opt2Nm			= Trim(SubNodes.getElementsByTagName("Opt2Nm").item(0).text)			'�ɼ��ڵ�2
				Opt2Tval		= Trim(SubNodes.getElementsByTagName("Opt2Tval").item(0).text)			'�ɼǸ�2
				ItemSaleStatCd	= Trim(SubNodes.getElementsByTagName("ItemSaleStatCd").item(0).text)	'��ǰ�ǸŻ���(10:�ǸŻ���, 20:ǰ��, 30:�Ǹ�����)

				rw "��ǰ�ɼǹ�ȣ : " & ItemNo
				rw "�ɼ������� : " & InvQty
				rw "�ɼ��ڵ�1 : " & Opt1Nm
				rw "�ɼǸ�1 : " & Opt1Tval
				rw "�ɼ��ڵ�2 : " & Opt2Nm
				rw "�ɼǸ�2 : " & Opt2Tval
				rw "��ǰ�ǸŻ��� : " & ItemSaleStatCd
				rw "<br>"
			Next
response.end
	End If
End Function

''���û�ǰ ��ȸ
function CheckLotteItemStat(iitemid,byRef iErrStr, byRef iSalePrc, byref iGoodsNm)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
    Dim iRbody

    CheckLotteItemStat = ""
    iLotteItemID = getLotteItemIdByTenItemID(iitemid)

    strParam = "subscriptionId=" & lotteAuthNo & "&strGoodsNo="&iLotteItemID

    On Error Resume Next
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
'' rw "https://openapi.lotte.com/openapi/searchGoodsListOpenApiOther.lotte?"&strParam ''''���û�ǰ��ȸaLL
	objXML.Open "POST", lotteAPIURL & "/openapi/searchGoodsListOpenApiOther.lotte", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)

	If objXML.Status = "200" Then
		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
		''rw iRbody
		iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> �� �� ������. ��ǰ�� < > ����..
		iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
		iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")

		xmlDOM.LoadXML iRbody

		if Err<>0 then
			iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-ItemChk-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'��� �ڵ�

		GoodsViewCount = xmlDOM.getElementsByTagName("GoodsViewCount").item(0).text  ''�����

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF
    ''rw "GoodsViewCount="&GoodsViewCount
    ''rw "Err="&Err
    	if (GoodsViewCount="1") then
        	SaleStatCd = xmlDOM.getElementsByTagName("SaleStatCd").item(0).text
        	iSalePrc   = xmlDOM.getElementsByTagName("SalePrc").item(0).text
        	iGoodsNm   = xmlDOM.getElementsByTagName("GoodsNm").item(0).text
        	iGoodsNm   = replace(iGoodsNm,"@@amp@@","&")

        end if

		'// ���� ����
		if Err<>0 then
	        if (IsAutoScript) then
	            rw "["&iitemid & "]:"&iMessage
	        else
				iErrStr =  "["&iitemid & "]:"&iMessage
			end if
			Set objXML = Nothing
		    Set xmlDOM = Nothing
		    On Error Goto 0
		    Exit function
        else

        end if
		Set xmlDOM = Nothing

		if (GoodsViewCount<>"1") then
		    iErrStr ="�˻������ �����ϴ�."&iMessage
		elseif (SaleStatCd<>"0") then
		    CheckLotteItemStat = SaleStatCd
	    end if
	else
		iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-002]"
	end if
	Set objXML = Nothing
	On Error Goto 0

'	rw "SaleStatCd="&SaleStatCd
'	rw "GoodsViewCount="&GoodsViewCount
'	rw "iMessage="&iMessage
'	rw "iErrStr="&iErrStr
end function

function checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm) ''�Ե����� (��)�ǸŻ���, (��)�ǸŰ� ������Ʈ, ��ǰ�� �߰�
    dim sqlstr, LotteSellyn
    dim assignedRow : assignedRow=0
    dim pbuf : pbuf =""

    iLotteItemStat= Trim(iLotteItemStat)
    iLotteGoodsNm=Trim(iLotteGoodsNm)
    iLotteGoodsNm=Replace(iLotteGoodsNm,"&gt;",">")
    iLotteGoodsNm=Replace(iLotteGoodsNm,"&lt;","<")
    iLotteGoodsNm=Replace(iLotteGoodsNm,"&nbsp;"," ")
    iLotteGoodsNm=Replace(iLotteGoodsNm,"&amp;","&")

    if (iLotteItemStat="10") then
        LotteSellyn = "Y"
    elseif (iLotteItemStat="20") then
        LotteSellyn = "N"
    elseif (iLotteItemStat="30") then
        LotteSellyn = "X"
    end if

    sqlstr = "select (lotteStatCd+','+LotteSellyn+','+convert(varchar(10),convert(int,LottePrice)) )as pbuf"
    sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
    sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
    rsget.Open sqlstr,dbget,1
    if Not rsget.Eof then
        pbuf = rsget("pbuf")
    end if
    rsget.close()


    sqlstr = "Update R" & VbCRLF
    sqlstr = sqlstr & " SET LottePrice="&iLotteSalePrc & VbCRLF
    ''sqlstr = sqlstr & " ,lotteStatCd='"&iLotteItemStat&"'"  ''�ٸ�������. lotteStatCd(�������), iLotteItemStat(�ǸŻ���)
    IF (LotteSellyn<>"") then
        sqlstr = sqlstr & " ,LotteSellyn='"&LotteSellyn&"'"
    ENd IF
    sqlstr = sqlstr & " ,regitemname='"&html2db(iLotteGoodsNm)&"'"
    sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
    sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
    sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
    sqlstr = sqlstr & " and (LottePrice<>"&iLotteSalePrc&"" & VbCRLF
    ''sqlstr = sqlstr & "     or lotteStatCd<>'"&iLotteItemStat&"'" & VbCRLF
    sqlstr = sqlstr & "     or LotteSellyn<>'"&LotteSellyn&"'"& VbCRLF
    sqlstr = sqlstr & "     or isNULL(regitemname,'')<>'"&html2db(iLotteGoodsNm)&"'"& VbCRLF
    sqlstr = sqlstr & " )" & VbCRLF

    ''rw sqlstr
    dbget.Execute sqlstr,assignedRow

    if (assignedRow<1) then
        ''�ٸ��� ������ lastStatCheckDate �� ������Ʈ
        sqlstr = "Update R" & VbCRLF
        sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
        sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
        sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
        dbget.Execute sqlstr
    else
        ''�ٸ��� ������ �α�.
        CALL Fn_AcctFailLog(CMALLNAME,iitemid,iLotteItemStat&","&LotteSellyn&","&iLotteSalePrc&"::"&pbuf,"STAT_CHK")

    end if
end function

Function CheckLotteTmpItemChk(iitemid,byRef iErrStr, byRef iLotteGoodNo, byref iLotteStatCd)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteTmpID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLotteTmpItemChk = ""
	iLotteTmpID = getLotteTmpItemIdByTenItemID(iitemid)
	If iLotteTmpID = "���û�ǰ" Then
		CheckLotteTmpItemChk = "���û�ǰ"
	Else
		On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", lotteAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte?subscriptionId=" & lotteAuthNo & "&goods_req_no=" & iLotteTmpID, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

					If Err <> 0 Then
						iErrStr = "�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�."
					    Set objXML = Nothing
					    Set xmlDOM = Nothing
					    Exit Function
					End If

					GoodsViewCount 		= Trim(xmlDOM.getElementsByTagName("Result").item(0).text)

				    If Err <> 0 Then
					    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
					End If

					If (GoodsViewCount = "1") Then
						iLotteGoodNo		= Trim(xmlDOM.getElementsByTagName("goods_no").item(0).text)			'���û�ǰ��ȣ
						iLotteStatCd		= Trim(xmlDOM.getElementsByTagName("conf_stat_cd").item(0).text)		'���������ڵ�
				    End If

					If Err <> 0 Then
						iErrStr =  "["&iitemid & "]:"&iMessage
						Set objXML = Nothing
					    Set xmlDOM = Nothing
					    On Error Goto 0
					    Exit Function
				    Else

				    End If
				Set xmlDOM = Nothing

				If (GoodsViewCount <> "1") Then
				    iErrStr ="�˻������ �����ϴ�."&iMessage
				Else
				    CheckLotteTmpItemChk = iLotteStatCd
			    End If
			Else
				iErrStr = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�..[ERR-ItemChk-002]"
			End If
		Set objXML = Nothing
		On Error Goto 0
	End If
End Function

function getTenItem2LotteGoodNo(iitemid)
    dim ret, sqlstr
    sqlstr = "select isNULL(isNULL(lotteGoodNo,lotteTmpGoodNo),'') as LGoodNo from db_item.dbo.tbl_lotte_regItem where itemid="&Left(iitemid,10)
 ''rw  sqlstr
    rsget.Open sqlstr,dbget,1
    if Not rsget.Eof then
        ret = rsget("LGoodNo")
    end if
    rsget.close()

    getTenItem2LotteGoodNo = Trim(ret)
end function

function getItemOptionCount(iitemid)
    dim ret, sqlstr
    ret = 0
    sqlstr = "select count(*) as optCNT from db_item.dbo.tbl_item_option where itemid="&Left(iitemid,10)&" and isusing='Y'"
    rsget.Open sqlstr,dbget,1
    if Not rsget.Eof then
        ret = rsget("optCNT")
    end if
    rsget.close()

    getItemOptionCount = Trim(ret)
end function

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

function XMLFileSaveLotte(xmlStr,mode,iitemid)
    Dim fso,tFile
    Dim opath : opath = "/admin/etc/Lotte/xmlFiles/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
    Dim defaultPath : defaultPath = server.mappath(opath) + "\"
    Dim fileName : fileName = mode &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"

    CALL CheckFolderCreate(defaultPath)

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(defaultPath & FileName )
	tFile.Write(xmlStr)
	tFile.Close
	Set tFile = nothing
    Set fso = nothing
end function


	dim mode, strSql, i, actCnt, strRst
	dim oLotteitem, arrItemid, chgSellYn
	dim strParam
	dim LotteGoodNo
	dim strArg, j,k,pp
	dim retErrStr, ret1, iErrStr
	Dim iMessage
	dim AssignedRow
	dim ilottegoods_no, delitemid

	dim iitemid, oneProdInfo, ProdCount, SubNodes
	dim buf, buf2, GoodNo, ItemNo, OptDesc, DispYn, SaleStatCd, StockQty, bufOpt
	dim buf_item_list
	dim iLotteItemStat,iLotteSalePrc, ArrRows, iLotteGoodsNm
	Dim iLotteItemTmpChk,iLotteStatCd
	Dim CNT10,CNT20,CNT30,iLotteGoodNo, iItemName, pregitemname
    Dim retFlag

	mode = request("mode")
	arrItemid = request("cksel")
	chgSellYn = request("chgSellYn")
    delitemid = requestCheckvar(request("delitemid"),10)
    retFlag   = request("retFlag")

	actCnt = 0		'ó�� ��ǰ �Ǽ�

''on Error Resume Next

	Select Case mode
		'----------------------------------------------------------------------
		'// �Ե����� �̵�� ��ǰ �ϰ� ���(50��)
		'----------------------------------------------------------------------
		Case "RegAll"
			rw "�̻��"
			response.end

		'----------------------------------------------------------------------
		'// ���� ��ǰ �ϰ� ���(�ִ� 20��)
		'----------------------------------------------------------------------
		Case "RegSelect"
		''rw "��ǰ����������"
		''response.end
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ���û�ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 20
			oLotteitem.FRectItemID	=arrItemid
			oLotteitem.getLotteNotRegItemList

            if (oLotteitem.FResultCount<1) then

                arrItemid = split(arrItemid,",")

                for i=LBound(arrItemid) to UBound(arrItemid)
                    CALL Fn_AcctFailTouch("lotteCom",arrItemid(i),"��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...")
                Next

                if (IsAutoScript) then
                    rw "S_ERR|��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, ����..."
                    dbget.Close: Response.End
                ELSE
                    Response.Write "<script language=javascript>alert('��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...');</script>"
    				dbget.Close: Response.End
    			ENd If
			end if

			for i=0 to (oLotteitem.FResultCount-1)
				''2012/09/10 �߰�
				strSql = "IF Exists(select * from db_item.dbo.tbl_lotte_regItem where itemid="&oLotteitem.FItemList(i).Fitemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " update R" & VbCRLF
    			strSql = strSql & "	Set LotteLastUpdate=getdate() "  & VbCRLF
        		strSql = strSql & "	, lottestatCD='10'"& VbCRLF                     ''����� ��Ͻõ��� �ϴ� ����(�ߺ���� �ȵǰ�)
    			strSql = strSql & "	From db_item.dbo.tbl_lotte_regItem R"& VbCRLF
    			strSql = strSql & " Where R.itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
    			strSql = strSql & " END ELSE "
    			strSql = strSql & " BEGIN"& VbCRLF
    			strSql = strSql & " Insert into db_item.dbo.tbl_lotte_regItem"
                strSql = strSql & " (itemid,regdate,reguserid,LotteStatCd)"
                strSql = strSql & " values ("&oLotteitem.FItemList(i).Fitemid&",getdate(),'"&session("SSBctID")&"','10')"
    			strSql = strSql & " END "
			    dbget.Execute strSql

				'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
				if oLotteitem.FItemList(i).checkTenItemOptionValid then
				    On Error Resume Next
					'//��ǰ��� �Ķ����
					strParam = oLotteitem.FItemList(i).getLotteItemRegParameter(FALSE)
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "" & strParam
END IF
					if Err<>0 then
					    rw Err.Description
						Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
						dbget.Close: Response.End
					end if
	                on Error Goto 0


	                iErrStr = ""
                    ret1 = LotteOneItemReg(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr,oLotteitem.FItemList(i).FSellCash,oLotteitem.FItemList(i).getLotteSellYn)
                    if (ret1) then
                        actCnt = actCnt+1
                    else
                        CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)

                        retErrStr = retErrStr & iErrStr
                    end if

				else
					'�ɼ� �˻翡 ������ ��ǰ�� ������ܻ�ǰ���� ���
'					strSql = "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'lotte' AND itemid = '" & oLotteitem.FItemList(i).Fitemid & "') " & _
'							"		BEGIN " & _
'							"			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun) VALUES('" & oLotteitem.FItemList(i).Fitemid & "','lotte') " & _
'							"		END	"
					''dbget.Execute(strSql)
					CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)

					iErrStr = "["&oLotteitem.FItemList(i).Fitemid&"] �ɼǰ˻� ����"
					retErrStr = retErrStr & iErrStr
				end if

			Next

			set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if

		'----------------------------------------------------------------------
		'// ������ ��ǰ �ϰ� ����(50��)
		'----------------------------------------------------------------------
		Case "EditAll"
		    response.write "������� param"
		    response.end

		'----------------------------------------------------------------------
		'// ���õ� ��ǰ �ϰ� ����(20��)
		'----------------------------------------------------------------------
		Case "EditSelect"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)
				'//��ǰ��� �Ķ����
				on Error Resume Next
				if (oLotteitem.FItemList(i).FmaySoldOut="Y") then
				    iErrStr = ""
				    chgSellYn = CHKIIF(oLotteitem.FItemList(i).FLotteSellYn="X","X","N")

                    if (LotteOneItemSellStatEdit(oLotteitem.FItemList(i).Fitemid,oLotteitem.FItemList(i).FLotteGoodNo,chgSellYn,iErrStr)) then
                        actCnt = actCnt+1
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&"ǰ��ó��"
                    else
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&iErrStr
                    end if
				else
'			2014-11-14 17:23 ������ // Ÿ�Ӿƿ��� �Ʒ� �����ΰ� �;� ��а� �ּ�ó�� ����
'					strParam = ""
'					strParam = oLotteitem.FItemList(i).getLotteCateParamToEdit()				'2013-07-30�����߰� // ����ī�װ� ����(��ǰ����API���� ���� �� ��)
'					If Err <> 0 Then
'						response.write Err.description
'						Response.Write "<script language=javascript>alert('�ٹ����� getLotteCateParamToEdit ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
'						dbget.Close: Response.End
'					End If
'
'					CALL LotteOneItemCateEdit(oLotteitem.FItemList(i).Fitemid, strParam, iErrStr)
'					rw iErrStr
'			2014-11-14 17:23 ������ // Ÿ�Ӿƿ��� �Ʒ� �����ΰ� �;� ��а� �ּ�ó�� ����
					strParam = ""
    				strParam = oLotteitem.FItemList(i).getLotteItemEditParameter()
    IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
    rw lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?" & strParam
    END IF
    				if Err<>0 then
    				    response.write Err.description
    					Response.Write "<script language=javascript>alert('�ٹ����� EditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
    					dbget.Close: Response.End
    				end if
                    on Error Goto 0

                    If (oLotteitem.FItemList(i).FoptionCnt>0) and (oLotteitem.FItemList(i).FregedOptCnt<1) Then
                        CALL LotteOneItemCheckStock(oLotteitem.FItemList(i).Fitemid,iErrStr)
                        rw iErrStr
                    End If

                    iErrStr = ""
                    ret1 = LotteOneItemInfoEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr,FALSE)

                    IF (ret1) THEN
                        '// ��ǰ���� ����
        				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
'        				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
        				''strSql = strSql & "	,accFailCnt=0"& VbCRLF  '' ���ݱ��� �Ѵ� �����Ǿ� 0ó�� // ���ݼ��� ������ ����.
        				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
        				dbget.Execute(strSql)
        				actCnt = actCnt+1
        			ELSE
        			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
        			    rw "[������������]"&iErrStr
                    ENd IF

           ''���� ���� 2011/11/11 �߰� eastone ------------------------------------------------------
                    IF  (ret1) and (oLotteitem.FItemList(i).FSellcash<>oLotteitem.FItemList(i).FLottePrice) THEN
                        strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()

        				if Err<>0 then
        					Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
        					dbget.Close: Response.End
        				end if

                        ret1 = false
                        ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

        				IF (ret1) THEN
        				    '// ��ǰ�������� ����
                			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
                			strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
                			strSql = strSql & "	, LottePrice=" & oLotteitem.FItemList(i).MustPrice & VbCRLF
                			strSql = strSql & "	, accFailCnt=0"& VbCRLF
                			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF

                			dbget.Execute(strSql)
                        ELSE
                            CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                            rw "[���ݼ�������]"&iErrStr
        				END IF
                    END IF
                    '2013-10-10 16:28 ������ �ϴ� �߰�(�����ȸ�κ�) �ѹ� �� ��ȸ
                    '2014-09-03 10:30 ������ �ϴ� IF�� �ּ�ó��
                    'If (oLotteitem.FItemList(i).FoptionCnt>0) and (oLotteitem.FItemList(i).FregedOptCnt<1) Then
						iErrStr = ""
						CALL LotteOneItemCheckStock(oLotteitem.FItemList(i).Fitemid,iErrStr)
					'End If


	                if (oLotteitem.FItemList(i).Fitemid<>"") then
	                    iLotteItemStat = CheckLotteItemStat(oLotteitem.FItemList(i).Fitemid,iErrStr,iLotteSalePrc,iLotteGoodsNm)
	
	                    if (iLotteItemStat<>"") then
	                        ''if (iLotteItemStat="10") then
	                            rw "["&oLotteitem.FItemList(i).Fitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
	                        ''end if
	
	                        CALL checkConfirmMatch(oLotteitem.FItemList(i).Fitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
	                    else
	                        rw "["&oLotteitem.FItemList(i).Fitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
	
	                        ''��� �ݺ��ϹǷ� �Ʒ� �ڵ� ����
	                        strSql = "Update R" & VbCRLF
	                        strSql = strSql & " SET lastStatCheckDate=getdate()"
	                        strSql = strSql & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
	                        strSql = strSql & " where R.itemid="&oLotteitem.FItemList(i).Fitemid & VbCRLF
	                   ' rw strSql
	                        dbget.Execute strSql,assignedRow
	                    end if
	                end if


                END IF

                retErrStr = retErrStr & iErrStr

                ''rw "Err.Number"&Err.Number
			Next

			set oLotteitem = Nothing
			''rw "Err.Number"&Err.Number
            ''rw "actCnt="&actCnt
            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if
            ''rw "Err.Number"&Err.Number

            if (retFlag<>"") then
                Response.Write "<script language=javascript>parent."&retFlag&";</script>"
                response.end
            end if
IF (session("ssBctID")="icommang") then
response.end
ENd IF
        Case "EditSelect2"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

            if (oLotteitem.FResultCount<1) then
               rw "���� ���ɻ�ǰ ����"& arrItemid
            end if

			for i=0 to (oLotteitem.FResultCount-1)
				'//��ǰ��� �Ķ����
				on Error Resume Next
				strParam = oLotteitem.FItemList(i).getLotteItemEditParameter2()
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "/openapi/upateApiNewGoodsInfo.lotte?" & strParam
END IF
				if Err<>0 then
				    response.write Err.description
					Response.Write "<script language=javascript>alert('�ٹ����� EditParameter2 ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				end if
                on Error Goto 0

'                IF (oLotteitem.FItemList(i).FoptionCnt>0) and (oLotteitem.FItemList(i).FregedOptCnt<1) then
'                    CALL LotteOneItemCheckStock(oLotteitem.FItemList(i).Fitemid,iErrStr)
'                    rw iErrStr
'                ENd If

                iErrStr = ""
                ret1 = LotteOneItemInfoEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr,TRUE)

                IF (ret1) THEN
                    '// ��ǰ���� ����
    				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
    				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
    				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
    				strSql = strSql & "	,accFailCnt=0"& VbCRLF
    				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
    				dbget.Execute(strSql)
    				actCnt = actCnt+1
    			ELSE
    			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
    			    rw "[������������]"&iErrStr
                ENd IF

                retErrStr = retErrStr & iErrStr

                ''rw "Err.Number"&Err.Number
			Next

			set oLotteitem = Nothing
			''rw "Err.Number"&Err.Number
            ''rw "actCnt="&actCnt
            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if
            ''rw "Err.Number"&Err.Number
IF (session("ssBctID")="icommang") then
response.end
ENd IF

		Case "EditSelect3"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)
				on Error Resume Next

				if (oLotteitem.FItemList(i).FmaySoldOut="Y") then
				    iErrStr = ""
				    chgSellYn = CHKIIF(oLotteitem.FItemList(i).FLotteSellYn="X","X","N")

                    if (LotteOneItemSellStatEdit(oLotteitem.FItemList(i).Fitemid,oLotteitem.FItemList(i).FLotteGoodNo,chgSellYn,iErrStr)) then
                        actCnt = actCnt+1
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&"ǰ��ó��"
                    else
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&iErrStr
                    end if
				else
					strParam = ""
					strParam = oLotteitem.FItemList(i).getLotteCateParamToEdit()				'2013-07-30�����߰� // ����ī�װ� ����(��ǰ����API���� ���� �� ��)
					If Err <> 0 Then
						response.write Err.description
						Response.Write "<script language=javascript>alert('�ٹ����� getLotteCateParamToEdit ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
						dbget.Close: Response.End
					End If

					CALL LotteOneItemCateEdit(oLotteitem.FItemList(i).Fitemid, strParam, iErrStr)
					rw iErrStr

					strParam = ""
    				strParam = oLotteitem.FItemList(i).getLotteItemEditParameter()
    IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
    rw lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?" & strParam
    END IF
    				if Err<>0 then
    				    response.write Err.description
    					Response.Write "<script language=javascript>alert('�ٹ����� EditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
    					dbget.Close: Response.End
    				end if
                    on Error Goto 0

                    If (oLotteitem.FItemList(i).FoptionCnt>0) and (oLotteitem.FItemList(i).FregedOptCnt<1) Then
                        CALL LotteOneItemCheckStock(oLotteitem.FItemList(i).Fitemid,iErrStr)
                        rw iErrStr
                    End If

                    iErrStr = ""
                    ret1 = LotteOneItemInfoEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr,FALSE)

                    IF (ret1) THEN
                        '// ��ǰ���� ����
        				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
        				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
        				''strSql = strSql & "	,accFailCnt=0"& VbCRLF  '' ���ݱ��� �Ѵ� �����Ǿ� 0ó�� // ���ݼ��� ������ ����.
        				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
        				dbget.Execute(strSql)
        				actCnt = actCnt+1
        			ELSE
        			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
        			    rw "[������������]"&iErrStr
                    ENd IF

           ''���� ���� 2011/11/11 �߰� eastone ------------------------------------------------------
                    IF  (ret1) and (oLotteitem.FItemList(i).FSellcash<>oLotteitem.FItemList(i).FLottePrice) THEN
                        strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()

        				if Err<>0 then
        					Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
        					dbget.Close: Response.End
        				end if

                        ret1 = false
                        ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

        				IF (ret1) THEN
        				    '// ��ǰ�������� ����
                			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
                			strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
                			strSql = strSql & "	, LottePrice=" & oLotteitem.FItemList(i).MustPrice & VbCRLF
                			strSql = strSql & "	, accFailCnt=0"& VbCRLF
                			strSql = strSql & "	, itemTableUpdateChkdate=getdate() "& VbCRLF
                			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF

                			dbget.Execute(strSql)
                        ELSE
                            CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                            rw "[���ݼ�������]"&iErrStr
        				END IF
                    END IF
                    '2013-10-10 16:28 ������ �ϴ� �߰�(�����ȸ�κ�) �ѹ� �� ��ȸ
                    If (oLotteitem.FItemList(i).FoptionCnt>0) and (oLotteitem.FItemList(i).FregedOptCnt<1) Then
						iErrStr = ""
						CALL LotteOneItemCheckStock(oLotteitem.FItemList(i).Fitemid,iErrStr)
					End If
                END IF

                retErrStr = retErrStr & iErrStr

                ''rw "Err.Number"&Err.Number
			Next

			set oLotteitem = Nothing
			''rw "Err.Number"&Err.Number
            ''rw "actCnt="&actCnt
            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if
            ''rw "Err.Number"&Err.Number

            if (retFlag<>"") then
                Response.Write "<script language=javascript>parent."&retFlag&";</script>"
                response.end
            end if
IF (session("ssBctID")="icommang") then
response.end
ENd IF

        '----------------------------------------------------------------------
		'// ���õ� ��ǰ ���� �ϰ� ����(20��) //�߰�
		'----------------------------------------------------------------------
		Case "EditPriceSelect"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)

       ''���� ���� 2011/11/11 �߰� eastone ------------------------------------------------------
                on Error Resume Next
                strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte?" & strParam
END IF
				if Err<>0 then
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				end if
                on Error Goto 0

                iErrStr = ""
				ret1 = false
                ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

				IF (ret1) THEN
				    '// ��ǰ�������� ����
        			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        			strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
        			strSql = strSql & "	, LottePrice=" & oLotteitem.FItemList(i).MustPrice & VbCRLF
        			strSql = strSql & "	, accFailCnt=0"& VbCRLF
        			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF
        			dbget.Execute(strSql)

        			actCnt = actCnt+1
                ELSE
                    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                    rw "iErrStr="&iErrStr
				END IF

                retErrStr = retErrStr & iErrStr
			Next

			set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if

        ''��ǰ�� ���� 2013/03/19 �߰� eastone ------------------------------------------------------
        Case "EditItemNm"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			end if

			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)

                on Error Resume Next
                strParam = oLotteitem.FItemList(i).getLotteItemNameEditParameter()
    IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
    rw lotteAPIURL & "/openapi/updateGoodsNmOpenApi.lotte?" & strParam
    END IF

    			if Err<>0 then
    				Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
    				dbget.Close: Response.End
    			end if
                on Error Goto 0

                iErrStr = ""
    			ret1 = false
                ret1 = LotteOneItemNameEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

    			IF (ret1) THEN
    			    '// ��ǰ�������� ����
        			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        			strSql = strSql & "	Set regitemname='" & html2db(oLotteitem.FItemList(i).FItemName) &"'"& VbCRLF
        			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF
        			dbget.Execute(strSql)

        			actCnt = actCnt+1
                ELSE
                    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                    rw "iErrStr="&iErrStr
    			END IF

                retErrStr = retErrStr & iErrStr
    		Next

    		set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if

		'----------------------------------------------------------------------
		'// ���û�ǰ �Ǹſ��� ����
		'// 2014-01-16 11:15 ������ �ǸŻ��¼����� ���� �ش��ǰ �ǸŻ���Ȯ�� �� �Ǹ������ ǰ�� OR �Ǹ������� ���� �� �ǰ� ����
		'----------------------------------------------------------------------
		Case "EditSellYn"
			Dim chkStat
			'## ������ ��ǰ ��� ����
			set oLotteitem = new CLotte
			oLotteitem.FPageSize	= 30
			oLotteitem.FRectItemID	= arrItemid
			''oLotteitem.FRectMatchCateNotCheck="on"
			oLotteitem.getLotteEditedItemList
			if (chgSellYn="N") and (oLotteitem.FResultCount<1) and (arrItemid="") then
			    oLotteitem.getLottereqExpireItemList
			end if

			rw oLotteitem.FResultCount

			for i=0 to (oLotteitem.FResultCount-1)
				chkStat = ""
				chkStat = CheckLotteItemStat(oLotteitem.FItemList(i).Fitemid,iErrStr,iLotteSalePrc,iLotteGoodsNm)
				If chkStat = "30" Then
					rw "["&oLotteitem.FItemList(i).Fitemid&"]"&" : �Ǹ������ �����Ұ�"
				Else
					'2014-02-06 12:00 ������ ���� �츮�� �ǸŻ��°� ǰ���̸� ������ N���� ����
					If (oLotteitem.FItemList(i).IsSoldOut) AND (chgSellYn <> "X") Then
						chgSellYn = "N"
					End If

				    iErrStr = ""
				    if (LotteOneItemSellStatEdit(oLotteitem.FItemList(i).Fitemid,oLotteitem.FItemList(i).FLotteGoodNo,chgSellYn,iErrStr)) then
				        actCnt = actCnt+1
				    else
				        rw "["&iitemid&"]"&iErrStr
				    end if
				    retErrStr = retErrStr & iErrStr
				End If
			Next

			set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if
		'----------------------------------------------------------------------
		'// ���޸� �ƴ� ��ǰ �ϰ� ����(20��)
		'----------------------------------------------------------------------
		Case "DelJaeHyu"
			'## ������ ��ǰ ��� ����
			rw "������� �޴�"
			response.end

			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectNotJehyu	= "Y"
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)
				'//��ǰ��� �Ķ����
				strParam = "?subscriptionId=" & lotteAuthNo											'�Ե����� ������ȣ	(*)
				strParam = strParam & "&goods_no=" & oLotteitem.FItemList(i).FLotteGoodNo			'�Ե����� ��ǰ��ȣ	(*)
				strParam = strParam & "&sale_stat_cd=20"											'�Ǹſ���(�Ǹ�����-�����Ұ�)

				Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.Open "GET", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam, false
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.Send()
				If objXML.Status = "200" Then

					'XML�� ���� DOM ��ü ����
					Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

					if Err<>0 then
						Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
						dbget.Close: Response.End
					end if

					'��� �ڵ�
					strRst = xmlDOM.getElementsByTagName("Result").item(0).text

					'// ���� ����
					if Err<>0 then
						Response.Write "<script language=javascript>alert('Error: " & xmlDOM.getElementsByTagName("Message").item(0).text  & "');</script>"
						dbget.Close: Response.End
					end if

					'// ��ǰ���� ����
					strSql = "Delete From db_item.dbo.tbl_lotte_regItem " &_
						" Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
					dbget.Execute(strSql)

					actCnt = actCnt+1

					Set xmlDOM = Nothing
				else
					Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
					dbget.Close: Response.End
				end if
				Set objXML = Nothing

			Next

			set oLotteitem = Nothing

		'----------------------------------------------------------------------
		'// ��ǰ ��� Ȯ��.
		'----------------------------------------------------------------------
	    Case "ChkStockSelect"
	        arrItemid = Trim(arrItemid)
	        if (Right(arrItemid,1)=",") then arrItemid=Left(arrItemid,Len(arrItemid)-1)
	        arrItemid = split(arrItemid,",")

	        for i=Lbound(arrItemid) to UBound(arrItemid)
	            iitemid        =arrItemid(i)

	            iErrStr = ""
				ret1 = false
                ret1 = LotteOneItemCheckStock(iitemid,iErrStr)

	            IF (ret1) THEN
        			actCnt = actCnt+1
                ELSE
                    rw "iErrStr="&iErrStr
				END IF

                retErrStr = retErrStr & iErrStr
	        Next

            if (retFlag<>"") then
                Response.Write "<script language=javascript>parent."&retFlag&";</script>"
                response.end
            end if
	        response.end
        Case "RegSelectWait"
            arrItemid = Trim(arrItemid)
            if Right(arrItemid,1)="," then arrItemid=Left(arrItemid,Len(arrItemid)-1)

            strSql = "Insert into db_item.dbo.tbl_lotte_regItem"
            strSql = strSql & " (itemid,regdate,reguserid,LotteStatCd)"
            strSql = strSql & " select i.itemid,getdate(),'"&session("SSBctID")&"','00'"
            strSql = strSql & " from db_item.dbo.tbl_item i"
            strSql = strSql & "     left join db_item.dbo.tbl_lotte_regItem R"
            strSql = strSql & "     on i.itemid=R.itemid"
            strSql = strSql & " where i.itemid in ("&arrItemid&")"
            strSql = strSql & " and R.itemid is NULL"
            dbget.Execute strSql,AssignedRow
 			response.write "<script>alert('"&AssignedRow&"�� ������ϵ�.');</script>"

			strSql = ""
			strSql = strSql & " update R "
			strSql = strSql & " set optAddPrcCnt= T.optAddPrcCnt"
			strSql = strSql & " from db_item.dbo.tbl_lotte_regitem R"
			strSql = strSql & " Join ("
			strSql = strSql & " 	select ii.itemid,count(*) as optAddPrcCnt"
			strSql = strSql & "		from db_item.dbo.tbl_item ii "
			strSql = strSql & " 	Join db_item.dbo.tbl_item_option o "
			strSql = strSql & " 	on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
			strSql = strSql & " 	group by ii.itemid "
			strSql = strSql & " ) T on R.itemid =T.itemid"
			strSql = strSql & " WHERE R.itemid in ("&arrItemid&")"
			dbget.Execute strSql,AssignedRow
			response.write "<script>parent.location.reload();</script>"
        Case "DelSelectWait"
            arrItemid = Trim(arrItemid)
            if Right(arrItemid,1)="," then arrItemid=Left(arrItemid,Len(arrItemid)-1)

            strSql = "delete from db_item.dbo.tbl_lotte_regItem"
            strSql = strSql & " where LotteStatCd in ('00')"
            strSql = strSql & " and itemid in ("&arrItemid&")"

            dbget.Execute strSql,AssignedRow
            response.write "<script>alert('"&AssignedRow&"�� ���� ������.');parent.location.reload();</script>"
        CASE "DelSelectExpireItem"
            ''API�� ���°� �Ǹ� �������� Ȯ�� �� ���� ��� �߰� �Ұ�
            iLotteItemStat = CheckLotteItemStat(delitemid,iErrStr,iLotteSalePrc,iLotteGoodsNm)
            if (iLotteItemStat="30") then
            	strSql = ""
                strSql = strSql & " delete from db_item.dbo.tbl_lotte_regItem"
                strSql = strSql & " where LotteSellyn in ('X')"
                strSql = strSql & " and itemid in ("&delitemid&")"
                dbget.Execute strSql,AssignedRow
				actCnt = AssignedRow

				strSql = ""
                strSql = strSql & " delete from db_item.dbo.tbl_Outmall_regedoption "
                strSql = strSql & " where itemid in ("&delitemid&")"
                strSql = strSql & " and mallid = '"&CMALLNAME&"' "
                dbget.Execute strSql
                
                response.write "<script>alert('"&AssignedRow&"��  ������.');</script>" ''parent.location.reload();
                dbget.Close() : response.end
            else
                if (iLotteItemStat="") and (iErrStr="�˻������ �����ϴ�.") then
                    strSql = "delete from db_item.dbo.tbl_lotte_regItem"
                    strSql = strSql & " where LotteSellyn in ('X')"
                    strSql = strSql & " and itemid in ("&delitemid&")"
                    dbget.Execute strSql,AssignedRow
                    actCnt = AssignedRow
                    response.write "<script>alert('���� - ERR: �ǸŻ��� : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
                else
                    response.write "<script>alert('ERR: �ǸŻ��� : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
                end if
            end if
        CASE "CheckItemStat"
            arrItemid = split(Trim(arrItemid),",")
            if IsArray(arrItemid) then
            for i=LBound(arrItemid) to UBound(arrItemid)
                iErrStr = ""
                iitemid = Trim(arrItemid(i))

                if (iitemid<>"") then
                    iLotteItemStat = CheckLotteItemStat(iitemid,iErrStr,iLotteSalePrc,iLotteGoodsNm)

                    if (iLotteItemStat<>"") then
                        ''if (iLotteItemStat="10") then
                            rw "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr
                        ''end if

                        CALL checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
                    else
                        rw "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr

                        ''��� �ݺ��ϹǷ� �Ʒ� �ڵ� ����
                        strSql = "Update R" & VbCRLF
                        strSql = strSql & " SET lastStatCheckDate=getdate()"
                        strSql = strSql & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
                        strSql = strSql & " where R.itemid="&iitemid & VbCRLF
                   ' rw strSql
                        dbget.Execute strSql,assignedRow
                    end if
                end if
            next
            end if

            response.end

		CASE "getconfirmList"
			Dim AnotherStat
			arrItemid = split(Trim(arrItemid),",")
			If IsArray(arrItemid) then
				For i=LBound(arrItemid) to UBound(arrItemid)
					iErrStr = ""
					iitemid = Trim(arrItemid(i))
					If (iitemid<>"") Then
						iLotteItemTmpChk = CheckLotteTmpItemChk(iitemid, iErrStr, iLotteGoodNo, iLotteStatCd)
						Select Case iLotteItemTmpChk
							Case "10"	AnotherStat = "�ӽõ��"
							Case "20"	AnotherStat = "���ο�û"
							Case "30"	AnotherStat = "���οϷ�"
							Case "40"	AnotherStat = "�ݷ�"
							Case "50"	AnotherStat = "���κҰ�"
							Case "51"	AnotherStat = "����ο�û"
							Case "52"	AnotherStat = "������û"
						End Select

						If (iLotteItemTmpChk = "30") Then
							strSql =""
							strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem "
							strSql = strSql & "	SET lastConfirmdate = getdate() "
							strSql = strSql & "	,LotteStatCd='30' "
							strSql = strSql & " ,LotteGoodNo='" & iLotteGoodNo & "' "
							strSql = strSql & " WHERE itemid='" & iitemid & "'"
							dbget.Execute strSql
							rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
						'2014-01-08 18:32 �������߰�//������ ��û..backyard �귣�� �űԻ�ǰ�� �ݷ�->�Ǹŷ� ������ ����, �Ե����ο��� �ϳ��� ���ο�û���� ���� �� �츮 ���ο��� ���Ȯ���� ���� �� statcd�����ϱ� ����
						ElseIf (iLotteItemTmpChk = "20") Then
							strSql =""
							strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem "
							strSql = strSql & "	SET LotteStatCd='20' "
							strSql = strSql & " WHERE itemid='" & iitemid & "'"
							dbget.Execute strSql
							rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
						ElseIf iLotteItemTmpChk = "���û�ǰ" Then
							rw "["&iitemid&"] :"&iLotteItemTmpChk&": �̹� ���� ��ǰ�Դϴ�"
						Else
						    rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
						End If
					End If
				Next
            End If
			response.end
'20130612 ���� �߰� (���û�ǰ����ȸ)
        CASE "StatWithOption"
	        arrItemid = Trim(arrItemid)
	        if (Right(arrItemid,1)=",") then arrItemid=Left(arrItemid,Len(arrItemid)-1)
	        arrItemid = split(arrItemid,",")

	        for i=Lbound(arrItemid) to UBound(arrItemid)
	            iitemid        =arrItemid(i)

	            iErrStr = ""
				ret1 = false
                ret1 = CheckLotteItemStatWithOption(iitemid,iErrStr)

	            IF (ret1) THEN
        			actCnt = actCnt+1
                ELSE
                    rw "iErrStr="&iErrStr
				END IF

                retErrStr = retErrStr & iErrStr
	        Next

            if (retFlag<>"") then
                Response.Write "<script language=javascript>parent."&retFlag&";</script>"
                response.end
            end if
	        response.end
        CASE "CheckItemNmAuto"
            buf = ""
            CNT10 = 0
            strSql = "select top 10 r.itemid,r.LotteGoodNo,i.ItemName"
            strSql = strSql & "	from db_item.dbo.tbl_lotte_regItem r"
            strSql = strSql & "		Join db_item.dbo.tbl_item i"
            strSql = strSql & "		on r.itemid=i.itemid"
            strSql = strSql & "	where r.regitemname is Not NULL"
            strSql = strSql & "	and r.regitemname<>i.itemname"
            strSql = strSql & "	and r.LotteGoodNo is Not NULL"
            strSql = strSql & "	order by r.lastStatCheckDate desc"
            ''strSql = strSql & "	order by LotteSellyn desc, lastStatCheckDate desc"

            rsget.Open strSql,dbget,1
            if not rsget.Eof then
                ArrRows = rsget.getRows()
            end if
            rsget.close

            if isArray(ArrRows) then

                For i =0 To UBound(ArrRows,2)
                    iErrStr = ""
                    iitemid = CStr(ArrRows(0,i))
                    buf = buf&iitemid&","
                    iLotteGoodNo = CStr(ArrRows(1,i))
                    iItemName    = CStr(ArrRows(2,i))

                    if (iitemid<>"") then
                        strParam = fnGetLotteItemNameEditParameter(iLotteGoodNo,iItemName)

            			ret1 = false
                        ret1 = LotteOneItemNameEdit(iitemid,strParam,iErrStr)

            			IF (ret1) THEN
            			    '// ��ǰ�� ����
            			    pregitemname = ""

            			    strSql = "select isNULL(regitemname,'') as regitemname from db_item.dbo.tbl_lotte_regItem "& VbCRLF
            			    strSql = strSql & "	Where itemid='" & iitemid & "'"& VbCRLF
            			    rsget.Open strSql,dbget,1
                            if not rsget.Eof then
                                pregitemname = rsget("regitemname")
                            end if
                            rsget.close

                			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
                			strSql = strSql & "	Set regitemname='" & html2db(iItemName) &"'"& VbCRLF
                			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
                			dbget.Execute(strSql)

                			CNT10 = CNT10+1

                			if (pregitemname<>iItemName) then
                			    buf2 = buf2 & pregitemname & "::" & iItemName &"<br>"
                			end if
                        ELSE
                            CALL Fn_AcctFailTouch("lotteCom",iitemid,iErrStr)
                            rw "iErrStr="&iErrStr
            			END IF
                    end if
                next
            end if
            rw buf
            rw buf2
            rw CNT10&"�� ��ǰ����� ����"
            response.end
		'2014-12-16 ������ �߰� (��ǰǰ�� ����)
		CASE "EditItemPO"
			If arrItemid = "" Then
				Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
				dbget.Close: Response.End
			End If

			'## ������ ��ǰ ��� ����
			Set oLotteitem = new CLotte
				oLotteitem.FPageSize       = 30
				oLotteitem.FRectItemID	= arrItemid
				oLotteitem.getLotteEditedItemList

			For i = 0 to (oLotteitem.FResultCount - 1)
                on Error Resume Next
                strParam = oLotteitem.FItemList(i).getLotteItemInfoCdToEdt()

				IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
					rw lotteAPIURL & "/openapi/upateApiDisplayGoodsItemInfo.lotte?" & strParam
				END IF

				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
                on Error Goto 0

                iErrStr = ""
    			ret1 = false
                ret1 = LotteOneItemPoomOkEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

    			IF (ret1) THEN
        			actCnt = actCnt+1
                ELSE
                    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                    rw "iErrStr="&iErrStr
    			END IF
                retErrStr = retErrStr & iErrStr
    		Next
    		set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if
        CASE "CheckItemStatAuto"
            buf = ""
            CNT10=0
            CNT20=0
            CNT30=0

            strSql = "select top 20 r.itemid "
            strSql = strSql & "	from db_item.dbo.tbl_lotte_regItem r"
            strSql = strSql & "	where LotteGoodno is Not NULL"
            strSql = strSql & "	order by r.lastStatCheckDate, (CASE WHEN r.lottesellyn='X' THEN '0' ELSE r.lottesellyn END), r.lotteLastUpdate , r.itemid desc"
            ''strSql = strSql & "	order by r.lastStatCheckDate, r.lotteLastUpdate, (CASE WHEN r.lottesellyn='X' THEN '0' ELSE r.lottesellyn END), r.itemid desc"

'            strSql = "select top 20 r.itemid "
'            strSql = strSql & "	from db_item.dbo.tbl_lotte_regItem r"
'            strSql = strSql & "	Join db_item.dbo.tbl_item i on r.itemid=i.itemid "
'            strSql = strSql & "	where r.LotteGoodno is Not NULL"
'            strSql = strSql & "	and i.lastupdate>r.lotteLastUpdate and i.lastupdate>'2012-01-01'"
'            strSql = strSql & "	order by r.lastStatCheckDate, (CASE WHEN r.lottesellyn='X' THEN '0' ELSE r.lottesellyn END), i.lastupdate , r.lotteLastUpdate , r.itemid desc"
'
'            strSql = "select top 20 r.itemid "
'            strSql = strSql & "	from db_item.dbo.tbl_lotte_regItem r"
'            strSql = strSql & "	Join db_item.dbo.tbl_item i on r.itemid=i.itemid "
'            strSql = strSql & "	where r.LotteGoodno is Not NULL"
'            strSql = strSql & "	order by (CASE WHEN i.lastupdate>r.lotteLastUpdate and r.lastStatCheckDate is NULL THEN NULL ELSE r.lastStatCheckDate end )"
'            strSql = strSql & "	,(CASE WHEN i.lastupdate>r.lotteLastUpdate then 0 else 1 end)"
'            strSql = strSql & "	, (CASE WHEN r.lottesellyn='X' THEN '0' ELSE r.lottesellyn END), i.lastupdate , r.lotteLastUpdate , r.itemid desc"


''rw strSql
            rsget.Open strSql,dbget,1
            if not rsget.Eof then
                ArrRows = rsget.getRows()
            end if
            rsget.close


            if isArray(ArrRows) then

                For i =0 To UBound(ArrRows,2)
                    iErrStr = ""
                    iitemid = CStr(ArrRows(0,i))
                    ''rw iitemid
                    if (iitemid<>"") then
                        iLotteItemStat = CheckLotteItemStat(iitemid,iErrStr,iLotteSalePrc,iLotteGoodsNm)

                        if (iLotteItemStat<>"") then
                            ''if (iLotteItemStat="10") then
                                buf=buf& "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr &"<br>"
                            ''end if
                            if (iLotteItemStat="10") then
                                CNT10=CNT10+1
                            elseif (iLotteItemStat="20") then
                                CNT20=CNT20+1
                            elseif (iLotteItemStat="30") then
                                CNT30=CNT30+1
                            end if

                            CALL checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm)
                        else
                            buf=buf& "["&iitemid&"] :"&iLotteItemStat&":"& iLotteSalePrc& ":"&iLotteGoodsNm&":"& iErrStr &"<br>"

                            ''��� �ݺ��ϹǷ� �Ʒ� �ڵ� ����
                            strSql = "Update R" & VbCRLF
                            strSql = strSql & " SET lastStatCheckDate=getdate()"
                            strSql = strSql & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
                            strSql = strSql & " where R.itemid="&iitemid & VbCRLF
                       ' rw strSql
                            dbget.Execute strSql,assignedRow
                        end if
                    end if
                next
            end if
            rw "STAT10:"&CNT10&"<br>STAT20:"&CNT20&"<br>STAT30:"&CNT30&"<br><br>"&buf
            response.end
        CASE "etcSongjangFin"
            strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=7"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
            strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
            'rw strSql
            dbget.Execute strSql,AssignedRow
            response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
		CASE "updateSendState"
			strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
			strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
			strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
			strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
			strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
			dbget.Execute strSql,AssignedRow
			response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
			response.end
	End Select

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
        if (IsAutoScript) then
            rw "OK|"& iMessage & "<br>"& actCnt & "���� ó���Ǿ����ϴ�."
        else
    	    Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "���� ó���Ǿ����ϴ�.');</script>"
        end if
    Else
        if (IsAutoScript) then
            rw "S_ERR|ó�� �߿� ������ �߻��߽��ϴ�"
        else
            Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
        end if
    End If

''on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->