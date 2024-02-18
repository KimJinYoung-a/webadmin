<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// 변수선언
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp, waitCnt, rjtCnt
    dim AssignedCNT
    dim param2 : param2=request("param2")

	actCnt = 0			'실갱신건수
    waitCnt = 0
	on Error Resume Next

	'// 롯데닷컴 임시등록 상품 조회(승인전,승인요청,재승인,수정요청 / 승인완료,반려,승인불가 제외)
	strSql = "Select top 100 itemid, LotteTmpGoodNo From db_item.dbo.tbl_lotte_regItem "
	strSql = strSql & " Where LotteStatCd in ('10','20','51','52') "
	'strSql = strSql & " Where LotteStatCd in ('10','51','52') "		'승인요청은 제외(첫등록:10, 승인요청상태:20 이므로, 전체 상품이 승인이 나지 않았을 경우 같은 상품을 계속 확인하게 되는 상황이 발생되어 강제로 10으로 변경 후 끝까지 돌릴 때에 사용)
	strSql = strSql & " and dateDiff(hh,IsNULL(lastConfirmdate,'2001-01-01'),getdate())>5"  ''승인요청상태에서 계속 요청하므로.
	strSql = strSql & " and LotteTmpGoodNo is Not NULL"
	''strSql = strSql & " and LotteTmpGoodNo='34561137'"
	strSql = strSql & " order by IsNULL(lastConfirmdate,'2001-01-01') asc, regdate asc" ''asc"  첫등록 부텀. LotteStatCd,

	if (param2="0") then
	    strSql = "Select top 100 r.itemid, r.LotteTmpGoodNo From db_item.dbo.tbl_lotte_regItem r"
	    strSql = strSql & "     Join db_item.dbo.tbl_item i"
	    strSql = strSql & "     on r.itemid=i.itemid"
	    strSql = strSql & "     and (i.sellyn<>'Y' or (i.sellcash>isNULL(lottePrice,0)))"
	    strSql = strSql & " Where r.LotteStatCd in ('20') " ''승인요청
	    strSql = strSql & " and dateDiff(hh,IsNULL(r.lastConfirmdate,'2001-01-01'),getdate())>5"  ''승인요청상태에서 계속 요청하므로.
	    strSql = strSql & " and r.LotteTmpGoodNo is Not NULL"
	    strSql = strSql & " order by IsNULL(r.lastConfirmdate,'2001-01-01') asc, r.regdate asc"
	end if

	rsget.Open strSql,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		'// 롯데닷컴 전시등록 상품 조회
		Do Until rsget.EOF
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", lotteAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte?subscriptionId=" & lotteAuthNo & "&goods_req_no=" & rsget("LotteTmpGoodNo"), false
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
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

				if Err<>0 then
					Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
					dbget.Close: Response.End
				end if

				LotteGoodNo		= Trim(xmlDOM.getElementsByTagName("goods_no").item(0).text)			'전시상품번호
				LotteStatCd		= Trim(xmlDOM.getElementsByTagName("conf_stat_cd").item(0).text)		'인증상태코드
''rw "LotteStatCd="&LotteStatCd
				'// 수정
				strSql = "Update db_item.dbo.tbl_lotte_regItem "
				strSql = strSql & "	Set lastConfirmdate=getdate()"
				if (LotteStatCd<>"") then
    				strSql = strSql & "	,LotteStatCd='" & LotteStatCd & "'"
    			end if

				if (LotteGoodNo>"0") and (LotteGoodNo<>"") then
					strSql = strSql & " , LotteGoodNo='" & LotteGoodNo & "'"
				end if

				strSql = strSql & " Where itemid='" & rsget("itemid") & "'"
				dbget.Execute strSql,AssignedCNT
				if (LotteStatCd="30") then
				    actCnt = actCnt+AssignedCNT
				elseif (LotteStatCd="20") then
				    waitCnt = waitCnt+AssignedCNT
				elseif (LotteStatCd="40") then
				    rjtCnt = rjtCnt+AssignedCNT
				end if
''rw AssignedCNT&":"&LotteStatCd&":"&LotteGoodNo
				Set xmlDOM = Nothing
			else
				Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
				dbget.Close: Response.End
			end if
			Set objXML = Nothing

			rsget.MoveNext
		Loop

	end if

	rsget.Close

	'##### DB 저장 처리 #####
    If Err.Number = 0 Then
    	if actCnt>0 or waitCnt>0 or rjtCnt>0 then
    	    if (IsAutoScript) then
    	        rw  "OK|"&actCnt & "건이 승인." & waitCnt& "건이 승인대기." & rjtCnt & "건 반려"
    	    else
    	        IF (session("ssBctID")="icommang") then
    	            rw actCnt & "건이 승인." & waitCnt& "건이 승인대기." & rjtCnt & "건 반려"
    	        else
        		    Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
        	    end if
        	end if
    	else
    	    if (IsAutoScript) then
    	        rw  "OK|"&actCnt & "건이 승인." & waitCnt& "건이 승인대기."  & rjtCnt & "건 반려"
    	    else
        		Response.Write "<script language=javascript>alert('갱신할 임시등록 상품이 없습니다.');parent.history.go(0);</script>"
        	end if
    	end if
    Else
        if (IsAutoScript) then
            rw "S_ERR|처리 중에 오류가 발생했습니다"
        else
            Response.Write "<script language=javascript>alert('처리 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
        end if
    End If

	on Error Goto 0


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->