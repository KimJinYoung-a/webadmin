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
''롯데닷컴 단품정보 조회 :: 판매여부등이 정확히 넘어오지 않는듯.., StockQty: 단품은 넘어오나, 복합옵션은..
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
        iErrStr = "["&iitemid&"] 롯데 닷컴 코드 없음."
        Exit function
	end if

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")

    strParam = "?subscriptionId=" & lotteAuthNo											'롯데닷컴 인증번호	(*)
    strParam = strParam & "&search_gubun=goods_no"
	strParam = strParam & "&search_text=" & ilottegoods_no		'롯뎃닷컴 상품번호	(*)

''rw lotteAPIURL & "/openapi/searchStockList.lotte" & strParam

	objXML.Open "GET", lotteAPIURL & "/openapi/searchStockList.lotte"&strParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

    ''rw "objXML.Status="&objXML.Status
	If objXML.Status = "200" Then

		'//전달받은 내용 확인

		buf = BinaryToText(objXML.ResponseBody, "euc-kr")

	''CALL XMLFileSaveLotte(buf,"STQ",ilottegoods_no)  ''slow

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML replace(buf,"&","＆")
'rw buf
		if Err<>0 then
		    iErrStr =  "롯데닷컴 결과 분석 중에 오류가 발생 [" & iitemid & "]:"
            Set objXML = Nothing
            Set xmlDOM = Nothing
		    Exit function

			''Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			''dbget.Close: Response.End
		end if

        ProdCount   = Trim(xmlDOM.getElementsByTagName("ProdCount").item(0).text)   '' 단품 갯수

       ''rw "ProdCount="&ProdCount
        IF (ProdCount="1") then         ''단품인경우 SKIP ==> 변경
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

            ''2013/05/30 추가
            strSql = " IF Exists(select * from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6)"
            strSql = strSql & " BEGIN"
            strSql = strSql & " DELETE from db_item.dbo.tbl_OutMall_regedoption where mallid='lotteCom' and itemid="&iitemid&" and Len(outmalloptCode)>6"
            strSql = strSql & " END"
            dbget.Execute strSql

            for each SubNodes in oneProdInfo
				GoodNo	    = Trim(SubNodes.getElementsByTagName("GoodNo").item(0).text)
                ItemNo	    = Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)        '' 단품코드 (숫자 0,1,2,)
                OptDesc	    = Trim(SubNodes.getElementsByTagName("OptDesc").item(0).text)
                DispYn	    = Trim(SubNodes.getElementsByTagName("DispYn").item(0).text)         ''N:안함 Y:전시
                SaleStatCd	= Trim(SubNodes.getElementsByTagName("SaleStatCd").item(0).text) ''판매진행, 판매종료, 품절'
                StockQty	= Trim(SubNodes.getElementsByTagName("StockQty").item(0).text)

                OptDesc = replace(OptDesc,"＆","&")
                if (SaleStatCd<>"판매진행") THEN
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
		            strSql = strSql & " ,'"&ItemNo&"'" ''임시로 롯데 코드 넣음 //2013/04/01
		            strSql = strSql & " ,'lotteCom'"
		            strSql = strSql & " ,'"&ItemNo&"'"
		            strSql = strSql & " ,'"&html2DB(OptDesc)&"'"
		            strSql = strSql & " ,'"&DispYn&"'"
			        strSql = strSql & " ,'Y'"
			        strSql = strSql & " ,"&StockQty
			        strSql = strSql & ")"

			        dbget.Execute strSql, AssignedRow
			     ''rw "AssignedRow="&AssignedRow
			        ''옵션 코드 매칭.
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
	    iErrStr =  "롯데닷컴과 통신중에 오류가 발생 [" & iitemid & "]:"
        Set objXML = Nothing
        Set xmlDOM = Nothing
	    Exit function

		''Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
		''dbget.Close: Response.End
	end if
    Set objXML = Nothing

    On Error Goto 0
end function

'전시카테고리 수정
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
				    iErrStr = "롯데아이몰 결과 분석 중에 오류가 발생했습니다.LotteOneItemCateEdit"
					dbget.Close: Exit Function
				End If

				'결과 코드
				on Error resume next
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
		        on Error Goto 0

				'// 오류 검출
				If Err <> 0 Then
		            iErrStr = "Error: " & xmlDOM.getElementsByTagName("Message").item(0).text
					dbget.Close: Exit Function
		        End If
			Set xmlDOM = Nothing
		Else
		    iErrStr ="롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요"
			dbget.Close: Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

''롯데닷컴 상품 등록
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
		'//전달받은 내용 확인
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End
        buf = BinaryToText(objXML.ResponseBody, "euc-kr")
		''CALL XMLFileSaveLotte(buf,"REG",oLotteitem.FItemList(i).FItemID)

		''유니코드로 저장됨..(옵션명 등)

		LotteGoodNo = ""
		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML buf ''BinaryToText(objXML.ResponseBody, "euc-kr")

		LotteGoodNo = xmlDOM.getElementsByTagName("goods_no").item(0).text

		'// 오류 검출(상품번호가 반드시 존재해야 됨)
		if LotteGoodNo="" then
		    iMessage = xmlDOM.getElementsByTagName("Message").item(0).text
		    iErrStr =  "상품 등록중 오류 [" & iitemid & "]:"&iMessage
            Set objXML = Nothing
            Set xmlDOM = Nothing
		    Exit function
		end if

		if Err<>0 then
			iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if


		'상품존재여부 확인
		strSql = "Select count(itemid) From db_item.dbo.tbl_lotte_regItem Where itemid='" & iitemid & "'"
		rsget.Open strSql,dbget,1

		if rsget(0)>0 then
			'// 존재 -> 수정
			strSql = "update R" & VbCRLF
			strSql = strSql & "	Set LotteLastUpdate=getdate() "  & VbCRLF
			strSql = strSql & "	, LotteTmpGoodNo='" & LotteGoodNo & "'"  & VbCRLF
			strSql = strSql & "	, LottePrice=" &iSellCash& VbCRLF
			strSql = strSql & "	, accFailCnt=0"& VbCRLF
			strSql = strSql & "	, lotteRegdate=isNULL(lotteRegdate,getdate())" ''추가 2013/02/26
			if (LotteGoodNo<>"") then
			    strSql = strSql & "	, lottestatCD='20'"& VbCRLF
			else
    			strSql = strSql & "	, lottestatCD='10'"& VbCRLF
    		end if
			strSql = strSql & "	From db_item.dbo.tbl_lotte_regItem R"& VbCRLF
			strSql = strSql & " Where R.itemid='" & iitemid & "'"

			dbget.Execute(strSql)
		else
			'// 없음 -> 신규등록
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
        ''옵션 리스트 저장(20120807 추가)
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
                    ''이중옵션인경우 이 구문 잘못 되었음.
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

			        ''옵션 코드 매칭.
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
		iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-REG-002]"
	end if
	Set objXML = Nothing
    On Error Goto 0

end function

''롯데닷컴 상품정보 수정
function LotteOneItemInfoEdit(iitemid,strParam,byRef iErrStr,isVer2)
    Dim objXML,xmlDOM,strRst, iMessage

    On Error Resume Next
    LotteOneItemInfoEdit = False

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    if (isVer2) then
        objXML.Open "POST", lotteAPIURL & "/openapi/upateApiNewGoodsInfo.lotte", false          ''상품수정
    else
        objXML.Open "POST", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte", false      ''전시상품수정
    end if

    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send(strParam)

	If objXML.Status = "200" Then

		'//전달받은 내용 확인
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
''response.end

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
		    iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
			''Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			''dbget.Close: Response.End
		end if

		'결과 코드

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

        if (strRst<>"1") then
            ''rw "FitemDiv="&oLotteitem.FItemList(i).FitemDiv  ''주문제작 상품인경우 수정이 안되는듯. add_choc_tp_cd_20 에 값이 있는경우;;
            ''rw "iMessage="&iMessage
            iErrStr =  "상품 수정중 오류 [" & iitemid & "]:"&iMessage
            Set objXML = Nothing
            Set xmlDOM = Nothing
            On Error Goto 0
		    Exit function
        end if

        if (strRst="1") then
			'// 오류 검출
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
	    iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-EDIT-002]"
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

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
			iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-PRCEDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'결과 코드

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

		'// 오류 검출
		if Err<>0 then
		    IF (Trim(iMessage)="002. 세일진행(또는 진행예정)") THEN
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
		iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-PRCEDIT-002]"
	end if
	Set objXML = Nothing
 On Error Goto 0
end function

function LotteOneItemSellStatEdit(iitemid,iLotteGoodNo,ichgSellYn,byRef iErrStr)
    Dim strParam
    Dim objXML, xmlDOM
    Dim strRst, strSql

    LotteOneItemSellStatEdit = False
	'//상품등록 파라메터
	strParam = "?subscriptionId=" & lotteAuthNo											'롯데닷컴 인증번호	(*)
	strParam = strParam & "&goods_no=" & iLotteGoodNo                       			'롯뎃닷컴 상품번호	(*)
	if ichgSellYn="Y" then																'판매여부(10:판매, 20:품절, 30:판매종료)
		strParam = strParam & "&sale_stat_cd=10"
	elseif ichgSellYn="N" then
		strParam = strParam & "&sale_stat_cd=20"
	elseif ichgSellYn="X" then                           '''X 기능 사용안함
		''''strParam = strParam & "&sale_stat_cd=30"      ''판매종료되면 수정못함.
		strParam = strParam & "&sale_stat_cd=20"
	end if

''rw lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
		    iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요"
			''Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			dbget.Close: Exit function
		end if

		'결과 코드
		on Error resume next
		strRst = xmlDOM.getElementsByTagName("Result").item(0).text
        on Error Goto 0

		'// 오류 검출
		if Err<>0 then
'           IF (xmlDOM.getElementsByTagName("Message").item(0).text="판매종료인 상품의 판매상태를 수정할 수 없습니다.") THEN
'''         '// 상품정보 수정 2012/03/06 ==> 주석제거 2013/02/25
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

			'// 상품정보 수정
			strSql = "Update db_item.dbo.tbl_lotte_regItem " & VbCRLF
			strSql = strSql & " Set LotteLastUpdate=getdate() " & VbCRLF
			strSql = strSql & " ,LotteSellYn='" & ichgSellYn & "'" & VbCRLF
			strSql = strSql & " Where itemid='" & iitemid & "'"
			dbget.Execute(strSql)
        end if


		Set xmlDOM = Nothing
	else
	    iErrStr ="롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요"
		''Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
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

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		if Err<>0 then
			iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-NMEDIT-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'결과 코드

		strRst = xmlDOM.getElementsByTagName("Result").item(0).text

        if Err<>0 then
    	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
    	ENd IF

		'// 오류 검출
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
		iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-NMEDIT-002]"
	end if
	Set objXML = Nothing
 On Error Goto 0
end function

'품목전송
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
					iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-PoomEDIT-001]"
					Set objXML = Nothing
					Set xmlDOM = Nothing
					Exit Function
				End If
				strRst = xmlDOM.getElementsByTagName("Result").item(0).text
				If strRst <> 1 Then
					iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
				Else
					rw "["&iitemid & "]:품목정보 수정 성공"
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
			iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-PoomEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
end function

''전시상품 상세조회
Function CheckLotteItemStatWithOption(iitemid,byRef iErrStr)
	Dim ilottegoods_no
	Dim objXML, xmlDOM
	Dim buf, oneProdInfo, strParam
	Dim ItemNo, InvQty, Opt1Nm, Opt1Tval, Opt2Nm, Opt2Tval, ItemSaleStatCd
	On Error Resume Next
	CheckLotteItemStatWithOption = False
	ilottegoods_no = getTenItem2LotteGoodNo(iitemid)
	If (ilottegoods_no = "") then
		iErrStr = "["&iitemid&"] 롯데 닷컴 코드 없음."
		Exit Function
	End If
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strParam = "?subscriptionId=" & lotteAuthNo											'롯데닷컴 인증번호	(*)
		strParam = strParam & "&strGoodsNo=" & ilottegoods_no
		objXML.Open "GET", lotteAPIURL & "/openapi/searchGoodsViewListOpenApi.lotte"&strParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

	If objXML.Status = "200" Then
		buf = BinaryToText(objXML.ResponseBody, "euc-kr")
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML replace(buf,"&","＆")
''rw buf
		If Err <> 0 Then
			iErrStr =  "롯데닷컴 결과 분석 중에 오류가 발생 [" & iitemid & "]:"
			Set objXML = Nothing
			Set xmlDOM = Nothing
			Exit function
		End If

			Set oneProdInfo = xmlDOM.getElementsByTagName("ItemInfo")
							'복합옵션으로 5개까지 되어있는 듯 하지만 10x10은 다중옵션이 2개라서 옵션코드,옵션명을 2까지만 변수에 담음
			For each SubNodes in oneProdInfo
				ItemNo			= Trim(SubNodes.getElementsByTagName("ItemNo").item(0).text)			'상품옵션번호
				InvQty			= Trim(SubNodes.getElementsByTagName("InvQty").item(0).text)			'옵션재고수량
				Opt1Nm 			= Trim(SubNodes.getElementsByTagName("Opt1Nm").item(0).text)			'옵션코드1
				Opt1Tval		= Trim(SubNodes.getElementsByTagName("Opt1Tval").item(0).text)			'옵션명1
				Opt2Nm			= Trim(SubNodes.getElementsByTagName("Opt2Nm").item(0).text)			'옵션코드2
				Opt2Tval		= Trim(SubNodes.getElementsByTagName("Opt2Tval").item(0).text)			'옵션명2
				ItemSaleStatCd	= Trim(SubNodes.getElementsByTagName("ItemSaleStatCd").item(0).text)	'단품판매상태(10:판매상태, 20:품절, 30:판매종료)

				rw "상품옵션번호 : " & ItemNo
				rw "옵션재고수량 : " & InvQty
				rw "옵션코드1 : " & Opt1Nm
				rw "옵션명1 : " & Opt1Tval
				rw "옵션코드2 : " & Opt2Nm
				rw "옵션명2 : " & Opt2Tval
				rw "단품판매상태 : " & ItemSaleStatCd
				rw "<br>"
			Next
response.end
	End If
End Function

''전시상품 조회
function CheckLotteItemStat(iitemid,byRef iErrStr, byRef iSalePrc, byref iGoodsNm)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim strParam, iLotteItemID , SaleStatCd, GoodsViewCount
    Dim iRbody

    CheckLotteItemStat = ""
    iLotteItemID = getLotteItemIdByTenItemID(iitemid)

    strParam = "subscriptionId=" & lotteAuthNo & "&strGoodsNo="&iLotteItemID

    On Error Resume Next
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
'' rw "https://openapi.lotte.com/openapi/searchGoodsListOpenApiOther.lotte?"&strParam ''''전시상품조회aLL
	objXML.Open "POST", lotteAPIURL & "/openapi/searchGoodsListOpenApiOther.lotte", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send(strParam)

	If objXML.Status = "200" Then
		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
		''rw iRbody
		iRbody = replace(iRbody,"&","@@amp@@")   '' <![CDATA[]]> 로 안 묶여옴. 상품명에 < > 있음..
		iRbody = replace(iRbody,"<GoodsNm>","<GoodsNm><![CDATA[")
		iRbody = replace(iRbody,"</GoodsNm>","]]></GoodsNm>")

		xmlDOM.LoadXML iRbody

		if Err<>0 then
			iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다.[ERR-ItemChk-001]"
		    Set objXML = Nothing
		    Set xmlDOM = Nothing
		    Exit function
		end if

		'결과 코드

		GoodsViewCount = xmlDOM.getElementsByTagName("GoodsViewCount").item(0).text  ''결과수

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

		'// 오류 검출
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
		    iErrStr ="검색결과가 없습니다."&iMessage
		elseif (SaleStatCd<>"0") then
		    CheckLotteItemStat = SaleStatCd
	    end if
	else
		iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-ItemChk-002]"
	end if
	Set objXML = Nothing
	On Error Goto 0

'	rw "SaleStatCd="&SaleStatCd
'	rw "GoodsViewCount="&GoodsViewCount
'	rw "iMessage="&iMessage
'	rw "iErrStr="&iErrStr
end function

function checkConfirmMatch(iitemid,iLotteItemStat,iLotteSalePrc,iLotteGoodsNm) ''롯데닷컴 (현)판매상태, (현)판매가 업데이트, 상품명 추가
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
    ''sqlstr = sqlstr & " ,lotteStatCd='"&iLotteItemStat&"'"  ''다른개념임. lotteStatCd(진행상태), iLotteItemStat(판매상태)
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
        ''다른게 없으면 lastStatCheckDate 만 업데이트
        sqlstr = "Update R" & VbCRLF
        sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
        sqlstr = sqlstr & " From db_item.dbo.tbl_lotte_regItem R" & VbCRLF
        sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
        dbget.Execute sqlstr
    else
        ''다른게 있으면 로그.
        CALL Fn_AcctFailLog(CMALLNAME,iitemid,iLotteItemStat&","&LotteSellyn&","&iLotteSalePrc&"::"&pbuf,"STAT_CHK")

    end if
end function

Function CheckLotteTmpItemChk(iitemid,byRef iErrStr, byRef iLotteGoodNo, byref iLotteStatCd)
	Dim objXML,xmlDOM,strRst,iMessage
	Dim strParam, iLotteTmpID , SaleStatCd, GoodsViewCount
	Dim iRbody

	CheckLotteTmpItemChk = ""
	iLotteTmpID = getLotteTmpItemIdByTenItemID(iitemid)
	If iLotteTmpID = "전시상품" Then
		CheckLotteTmpItemChk = "전시상품"
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
						iErrStr = "롯데닷컴 결과 분석 중에 오류가 발생했습니다."
					    Set objXML = Nothing
					    Set xmlDOM = Nothing
					    Exit Function
					End If

					GoodsViewCount 		= Trim(xmlDOM.getElementsByTagName("Result").item(0).text)

				    If Err <> 0 Then
					    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
					End If

					If (GoodsViewCount = "1") Then
						iLotteGoodNo		= Trim(xmlDOM.getElementsByTagName("goods_no").item(0).text)			'전시상품번호
						iLotteStatCd		= Trim(xmlDOM.getElementsByTagName("conf_stat_cd").item(0).text)		'인증상태코드
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
				    iErrStr ="검색결과가 없습니다."&iMessage
				Else
				    CheckLotteTmpItemChk = iLotteStatCd
			    End If
			Else
				iErrStr = "롯데닷컴과 통신중에 오류가 발생했습니다..[ERR-ItemChk-002]"
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

	actCnt = 0		'처리 상품 건수

''on Error Resume Next

	Select Case mode
		'----------------------------------------------------------------------
		'// 롯데닷컴 미등록 상품 일괄 등록(50건)
		'----------------------------------------------------------------------
		Case "RegAll"
			rw "미사용"
			response.end

		'----------------------------------------------------------------------
		'// 선택 상품 일괄 등록(최대 20건)
		'----------------------------------------------------------------------
		Case "RegSelect"
		''rw "상품등록잠시중지"
		''response.end
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 선택상품 목록 접수
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 20
			oLotteitem.FRectItemID	=arrItemid
			oLotteitem.getLotteNotRegItemList

            if (oLotteitem.FResultCount<1) then

                arrItemid = split(arrItemid,",")

                for i=LBound(arrItemid) to UBound(arrItemid)
                    CALL Fn_AcctFailTouch("lotteCom",arrItemid(i),"등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...")
                Next

                if (IsAutoScript) then
                    rw "S_ERR|등록가능상품 없음 :등록조건 확인: 판매Y, 할인..."
                    dbget.Close: Response.End
                ELSE
                    Response.Write "<script language=javascript>alert('등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...');</script>"
    				dbget.Close: Response.End
    			ENd If
			end if

			for i=0 to (oLotteitem.FResultCount-1)
				''2012/09/10 추가
				strSql = "IF Exists(select * from db_item.dbo.tbl_lotte_regItem where itemid="&oLotteitem.FItemList(i).Fitemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " update R" & VbCRLF
    			strSql = strSql & "	Set LotteLastUpdate=getdate() "  & VbCRLF
        		strSql = strSql & "	, lottestatCD='10'"& VbCRLF                     ''통신전 등록시도로 일단 넣음(중복등록 안되게)
    			strSql = strSql & "	From db_item.dbo.tbl_lotte_regItem R"& VbCRLF
    			strSql = strSql & " Where R.itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
    			strSql = strSql & " END ELSE "
    			strSql = strSql & " BEGIN"& VbCRLF
    			strSql = strSql & " Insert into db_item.dbo.tbl_lotte_regItem"
                strSql = strSql & " (itemid,regdate,reguserid,LotteStatCd)"
                strSql = strSql & " values ("&oLotteitem.FItemList(i).Fitemid&",getdate(),'"&session("SSBctID")&"','10')"
    			strSql = strSql & " END "
			    dbget.Execute strSql

				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				if oLotteitem.FItemList(i).checkTenItemOptionValid then
				    On Error Resume Next
					'//상품등록 파라메터
					strParam = oLotteitem.FItemList(i).getLotteItemRegParameter(FALSE)
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "" & strParam
END IF
					if Err<>0 then
					    rw Err.Description
						Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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
					'옵션 검사에 실패한 상품은 등록제외상품으로 등록
'					strSql = "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'lotte' AND itemid = '" & oLotteitem.FItemList(i).Fitemid & "') " & _
'							"		BEGIN " & _
'							"			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun) VALUES('" & oLotteitem.FItemList(i).Fitemid & "','lotte') " & _
'							"		END	"
					''dbget.Execute(strSql)
					CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)

					iErrStr = "["&oLotteitem.FItemList(i).Fitemid&"] 옵션검사 실패"
					retErrStr = retErrStr & iErrStr
				end if

			Next

			set oLotteitem = Nothing

            if (retErrStr<>"") then
                Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
            end if

		'----------------------------------------------------------------------
		'// 수정된 상품 일괄 수정(50건)
		'----------------------------------------------------------------------
		Case "EditAll"
		    response.write "사용중지 param"
		    response.end

		'----------------------------------------------------------------------
		'// 선택된 상품 일괄 수정(20건)
		'----------------------------------------------------------------------
		Case "EditSelect"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 수정된 상품 목록 접수
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)
				'//상품등록 파라메터
				on Error Resume Next
				if (oLotteitem.FItemList(i).FmaySoldOut="Y") then
				    iErrStr = ""
				    chgSellYn = CHKIIF(oLotteitem.FItemList(i).FLotteSellYn="X","X","N")

                    if (LotteOneItemSellStatEdit(oLotteitem.FItemList(i).Fitemid,oLotteitem.FItemList(i).FLotteGoodNo,chgSellYn,iErrStr)) then
                        actCnt = actCnt+1
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&"품절처리"
                    else
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&iErrStr
                    end if
				else
'			2014-11-14 17:23 김진영 // 타임아웃이 아래 원인인가 싶어 당분간 주석처리 시작
'					strParam = ""
'					strParam = oLotteitem.FItemList(i).getLotteCateParamToEdit()				'2013-07-30진영추가 // 전시카테고리 수정(상품수정API에서 수정 안 됨)
'					If Err <> 0 Then
'						response.write Err.description
'						Response.Write "<script language=javascript>alert('텐바이텐 getLotteCateParamToEdit 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
'						dbget.Close: Response.End
'					End If
'
'					CALL LotteOneItemCateEdit(oLotteitem.FItemList(i).Fitemid, strParam, iErrStr)
'					rw iErrStr
'			2014-11-14 17:23 김진영 // 타임아웃이 아래 원인인가 싶어 당분간 주석처리 시작
					strParam = ""
    				strParam = oLotteitem.FItemList(i).getLotteItemEditParameter()
    IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
    rw lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte?" & strParam
    END IF
    				if Err<>0 then
    				    response.write Err.description
    					Response.Write "<script language=javascript>alert('텐바이텐 EditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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
                        '// 상품정보 수정
        				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
'        				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
        				''strSql = strSql & "	,accFailCnt=0"& VbCRLF  '' 가격까지 둘다 수정되야 0처리 // 가격수정 오류로 인해.
        				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
        				dbget.Execute(strSql)
        				actCnt = actCnt+1
        			ELSE
        			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
        			    rw "[정보수정오류]"&iErrStr
                    ENd IF

           ''가격 수정 2011/11/11 추가 eastone ------------------------------------------------------
                    IF  (ret1) and (oLotteitem.FItemList(i).FSellcash<>oLotteitem.FItemList(i).FLottePrice) THEN
                        strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()

        				if Err<>0 then
        					Response.Write "<script language=javascript>alert('텐바이텐 PriceEditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
        					dbget.Close: Response.End
        				end if

                        ret1 = false
                        ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

        				IF (ret1) THEN
        				    '// 상품가격정보 수정
                			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
                			strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
                			strSql = strSql & "	, LottePrice=" & oLotteitem.FItemList(i).MustPrice & VbCRLF
                			strSql = strSql & "	, accFailCnt=0"& VbCRLF
                			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF

                			dbget.Execute(strSql)
                        ELSE
                            CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                            rw "[가격수정오류]"&iErrStr
        				END IF
                    END IF
                    '2013-10-10 16:28 김진영 하단 추가(재고조회부분) 한번 더 조회
                    '2014-09-03 10:30 김진영 하단 IF문 주석처리
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
	
	                        ''계속 반복하므로 아래 코드 넣음
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
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 수정된 상품 목록 접수
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

            if (oLotteitem.FResultCount<1) then
               rw "수정 가능상품 없음"& arrItemid
            end if

			for i=0 to (oLotteitem.FResultCount-1)
				'//상품등록 파라메터
				on Error Resume Next
				strParam = oLotteitem.FItemList(i).getLotteItemEditParameter2()
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "/openapi/upateApiNewGoodsInfo.lotte?" & strParam
END IF
				if Err<>0 then
				    response.write Err.description
					Response.Write "<script language=javascript>alert('텐바이텐 EditParameter2 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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
                    '// 상품정보 수정
    				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
    				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
    				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
    				strSql = strSql & "	,accFailCnt=0"& VbCRLF
    				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
    				dbget.Execute(strSql)
    				actCnt = actCnt+1
    			ELSE
    			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
    			    rw "[정보수정오류]"&iErrStr
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
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 수정된 상품 목록 접수
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
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&"품절처리"
                    else
                        rw "["&oLotteitem.FItemList(i).Fitemid&"]"&iErrStr
                    end if
				else
					strParam = ""
					strParam = oLotteitem.FItemList(i).getLotteCateParamToEdit()				'2013-07-30진영추가 // 전시카테고리 수정(상품수정API에서 수정 안 됨)
					If Err <> 0 Then
						response.write Err.description
						Response.Write "<script language=javascript>alert('텐바이텐 getLotteCateParamToEdit 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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
    					Response.Write "<script language=javascript>alert('텐바이텐 EditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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
                        '// 상품정보 수정
        				strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
        				strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
        				strSql = strSql & "	,LotteSellYn='" & oLotteitem.FItemList(i).getLotteSellYn & "'" & VbCRLF
        				''strSql = strSql & "	,accFailCnt=0"& VbCRLF  '' 가격까지 둘다 수정되야 0처리 // 가격수정 오류로 인해.
        				strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
        				dbget.Execute(strSql)
        				actCnt = actCnt+1
        			ELSE
        			    CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
        			    rw "[정보수정오류]"&iErrStr
                    ENd IF

           ''가격 수정 2011/11/11 추가 eastone ------------------------------------------------------
                    IF  (ret1) and (oLotteitem.FItemList(i).FSellcash<>oLotteitem.FItemList(i).FLottePrice) THEN
                        strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()

        				if Err<>0 then
        					Response.Write "<script language=javascript>alert('텐바이텐 PriceEditParameter 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
        					dbget.Close: Response.End
        				end if

                        ret1 = false
                        ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

        				IF (ret1) THEN
        				    '// 상품가격정보 수정
                			strSql = "update db_item.dbo.tbl_lotte_regItem  " & VbCRLF
                			strSql = strSql & "	Set LotteLastUpdate=getdate() " & VbCRLF
                			strSql = strSql & "	, LottePrice=" & oLotteitem.FItemList(i).MustPrice & VbCRLF
                			strSql = strSql & "	, accFailCnt=0"& VbCRLF
                			strSql = strSql & "	, itemTableUpdateChkdate=getdate() "& VbCRLF
                			strSql = strSql & " Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"& VbCRLF

                			dbget.Execute(strSql)
                        ELSE
                            CALL Fn_AcctFailTouch("lotteCom",oLotteitem.FItemList(i).Fitemid,iErrStr)
                            rw "[가격수정오류]"&iErrStr
        				END IF
                    END IF
                    '2013-10-10 16:28 김진영 하단 추가(재고조회부분) 한번 더 조회
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
		'// 선택된 상품 가격 일괄 수정(20건) //추가
		'----------------------------------------------------------------------
		Case "EditPriceSelect"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 수정된 상품 목록 접수
			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectItemID	= arrItemid
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)

       ''가격 수정 2011/11/11 추가 eastone ------------------------------------------------------
                on Error Resume Next
                strParam = oLotteitem.FItemList(i).getLotteItemPriceEditParameter()
IF (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then
rw lotteAPIURL & "/openapi/updateGoodsSalePrcOpenApi.lotte?" & strParam
END IF
				if Err<>0 then
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				end if
                on Error Goto 0

                iErrStr = ""
				ret1 = false
                ret1 = LotteOnItemPriceEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

				IF (ret1) THEN
				    '// 상품가격정보 수정
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

        ''상품명 수정 2013/03/19 추가 eastone ------------------------------------------------------
        Case "EditItemNm"
			if arrItemid="" then
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			end if

			'## 수정된 상품 목록 접수
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
    				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
    				dbget.Close: Response.End
    			end if
                on Error Goto 0

                iErrStr = ""
    			ret1 = false
                ret1 = LotteOneItemNameEdit(oLotteitem.FItemList(i).Fitemid,strParam,iErrStr)

    			IF (ret1) THEN
    			    '// 상품가격정보 수정
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
		'// 선택상품 판매여부 수정
		'// 2014-01-16 11:15 김진영 판매상태수정시 먼저 해당상품 판매상태확인 후 판매종료면 품절 OR 판매중으로 변경 안 되게 수정
		'----------------------------------------------------------------------
		Case "EditSellYn"
			Dim chkStat
			'## 수정된 상품 목록 접수
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
					rw "["&oLotteitem.FItemList(i).Fitemid&"]"&" : 판매종료라서 수정불가"
				Else
					'2014-02-06 12:00 김진영 수정 우리쪽 판매상태가 품절이면 무조건 N으로 전송
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
		'// 제휴몰 아닌 상품 일괄 삭제(20건)
		'----------------------------------------------------------------------
		Case "DelJaeHyu"
			'## 수정된 상품 목록 접수
			rw "사용중지 메뉴"
			response.end

			set oLotteitem = new CLotte
			oLotteitem.FPageSize       = 30
			oLotteitem.FRectNotJehyu	= "Y"
			oLotteitem.getLotteEditedItemList

			for i=0 to (oLotteitem.FResultCount-1)
				'//상품등록 파라메터
				strParam = "?subscriptionId=" & lotteAuthNo											'롯데닷컴 인증번호	(*)
				strParam = strParam & "&goods_no=" & oLotteitem.FItemList(i).FLotteGoodNo			'롯뎃닷컴 상품번호	(*)
				strParam = strParam & "&sale_stat_cd=20"											'판매여부(판매종료-복구불가)

				Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.Open "GET", lotteAPIURL & "/openapi/upateApiDisplayGoodsInfo.lotte" & strParam, false
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.Send()
				If objXML.Status = "200" Then

					'XML을 담을 DOM 객체 생성
					Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

					if Err<>0 then
						Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
						dbget.Close: Response.End
					end if

					'결과 코드
					strRst = xmlDOM.getElementsByTagName("Result").item(0).text

					'// 오류 검출
					if Err<>0 then
						Response.Write "<script language=javascript>alert('Error: " & xmlDOM.getElementsByTagName("Message").item(0).text  & "');</script>"
						dbget.Close: Response.End
					end if

					'// 상품정보 삭제
					strSql = "Delete From db_item.dbo.tbl_lotte_regItem " &_
						" Where itemid='" & oLotteitem.FItemList(i).Fitemid & "'"
					dbget.Execute(strSql)

					actCnt = actCnt+1

					Set xmlDOM = Nothing
				else
					Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
					dbget.Close: Response.End
				end if
				Set objXML = Nothing

			Next

			set oLotteitem = Nothing

		'----------------------------------------------------------------------
		'// 단품 재고 확인.
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
 			response.write "<script>alert('"&AssignedRow&"건 예정등록됨.');</script>"

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
            response.write "<script>alert('"&AssignedRow&"건 예정 삭제됨.');parent.location.reload();</script>"
        CASE "DelSelectExpireItem"
            ''API로 상태가 판매 종료인지 확인 후 삭제 모듈 추가 할것
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
                
                response.write "<script>alert('"&AssignedRow&"건  삭제됨.');</script>" ''parent.location.reload();
                dbget.Close() : response.end
            else
                if (iLotteItemStat="") and (iErrStr="검색결과가 없습니다.") then
                    strSql = "delete from db_item.dbo.tbl_lotte_regItem"
                    strSql = strSql & " where LotteSellyn in ('X')"
                    strSql = strSql & " and itemid in ("&delitemid&")"
                    dbget.Execute strSql,AssignedRow
                    actCnt = AssignedRow
                    response.write "<script>alert('삭제 - ERR: 판매상태 : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
                else
                    response.write "<script>alert('ERR: 판매상태 : ["&iLotteItemStat&"]:"&iErrStr&"');</script>"
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

                        ''계속 반복하므로 아래 코드 넣음
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
							Case "10"	AnotherStat = "임시등록"
							Case "20"	AnotherStat = "승인요청"
							Case "30"	AnotherStat = "승인완료"
							Case "40"	AnotherStat = "반려"
							Case "50"	AnotherStat = "승인불가"
							Case "51"	AnotherStat = "재승인요청"
							Case "52"	AnotherStat = "수정요청"
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
						'2014-01-08 18:32 김진영추가//유미희 요청..backyard 브랜드 신규상품이 반려->판매로 돌리기 원함, 롯데어드민에서 하나씩 승인요청으로 변경 후 우리 어드민에서 등록확인을 했을 때 statcd변경하기 위함
						ElseIf (iLotteItemTmpChk = "20") Then
							strSql =""
							strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_regItem "
							strSql = strSql & "	SET LotteStatCd='20' "
							strSql = strSql & " WHERE itemid='" & iitemid & "'"
							dbget.Execute strSql
							rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
						ElseIf iLotteItemTmpChk = "전시상품" Then
							rw "["&iitemid&"] :"&iLotteItemTmpChk&": 이미 전시 상품입니다"
						Else
						    rw "["&iitemid&"] :"&iLotteItemTmpChk&":"&AnotherStat
						End If
					End If
				Next
            End If
			response.end
'20130612 진영 추가 (전시상품상세조회)
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
            			    '// 상품명 수정
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
            rw CNT10&"건 상품명수정 성공"
            response.end
		'2014-12-16 김진영 추가 (상품품목 수정)
		CASE "EditItemPO"
			If arrItemid = "" Then
				Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
				dbget.Close: Response.End
			End If

			'## 수정된 상품 목록 접수
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
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oLotteitem.FItemList(i).Fitemid & "]');</script>"
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

                            ''계속 반복하므로 아래 코드 넣음
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
            response.write "<script>alert('"&AssignedRow&"건 완료 처리.');opener.close();window.close()</script>"
		CASE "updateSendState"
			strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
			strSql = strSql & "	Set sendState='"&request("updateSendState")&"'"
			strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
			strSql = strSql & "	where OutMallOrderSerial='"&request("ORG_ord_no")&"'"
			strSql = strSql & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
			dbget.Execute strSql,AssignedRow
			response.write "<script>alert('"&AssignedRow&"건 완료 처리.');opener.close();window.close()</script>"
			response.end
	End Select

	'##### DB 저장 처리 #####
    If Err.Number = 0 Then
        if (IsAutoScript) then
            rw "OK|"& iMessage & "<br>"& actCnt & "건이 처리되었습니다."
        else
    	    Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "건이 처리되었습니다.');</script>"
        end if
    Else
        if (IsAutoScript) then
            rw "S_ERR|처리 중에 오류가 발생했습니다"
        else
            Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
        end if
    End If

''on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->