<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
Dim iURI : iURI = "/openapi/searchReturnListOpenApi.lotte"
Dim ret_no
Dim paramname

rw iURI
rw ret_no
rw lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo & "&ret_no=" & ret_no
Dim retText, retVal, retCnt
Dim i
Dim ReturnInfo, SubNodes, GoodsArtc, ArtcItem, SubSubNodes
Dim ReturnCode,ReturnName,ReturnAddress
Dim sqlStr

Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo & "&ret_no=" & ret_no, false  ''ret_no 안먹음..
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then
	    sqlStr = "update db_temp.dbo.tbl_jaehyumall_returnInfo"
	    sqlStr = sqlStr & " set isusing='N'"
	    sqlStr = sqlStr & " where mallgubun='lotteCom'"
	    sqlStr = sqlStr & " and  isusing='Y'"
	    dbget.Execute sqlStr

	    'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
		retCnt = xmlDOM.getElementsByTagName("ReturnCount").item(0).text		'결과수

		if (retVal) and (retCnt>0) then
            Set ReturnInfo = xmlDOM.getElementsByTagName("ReturnInfo")
		    for each SubNodes in ReturnInfo
		        ReturnCode		= Trim(SubNodes.getElementsByTagName("ReturnCode").item(0).text)		'배송지일련번호
		        ReturnName      = Trim(SubNodes.getElementsByTagName("ReturnName").item(0).text)		'배송지명
		        ReturnAddress   = Trim(SubNodes.getElementsByTagName("ReturnAddress").item(0).text)		'우편번호,기본주소,상세주소
		        '담당자
		        '연락처
		        '휴대폰
		        if (ReturnCode<>"72125") and (ReturnCode<>"") then ''기본배송지.
		            sqlStr = "if Exists(select * from db_temp.dbo.tbl_jaehyumall_returnInfo where mallgubun='lotteCom' and ReturnCode='"&ReturnCode&"') "&VbCRLF
		            sqlStr = sqlStr & " BEGIN"&VbCRLF
		            sqlStr = sqlStr & "  update db_temp.dbo.tbl_jaehyumall_returnInfo"&VbCRLF
		            sqlStr = sqlStr & "  set ReturnName='"&html2db(ReturnName)&"'"&VbCRLF
		            sqlStr = sqlStr & "  , ReturnAddress='"&html2db(ReturnAddress)&"'"&VbCRLF
		            sqlStr = sqlStr & "  , isUsing='Y'"&VbCRLF
		            sqlStr = sqlStr & "  , lastupdate=getdate()"&VbCRLF
		            sqlStr = sqlStr & "  where mallgubun='lotteCom'"&VbCRLF
		            sqlStr = sqlStr & "  and ReturnCode='"&ReturnCode&"'"&VbCRLF
		            sqlStr = sqlStr & " END"&VbCRLF
		            sqlStr = sqlStr & " ELSE "&VbCRLF
		            sqlStr = sqlStr & " BEGIN"&VbCRLF
		            sqlStr = sqlStr & "  Insert into db_temp.dbo.tbl_jaehyumall_returnInfo"&VbCRLF
		            sqlStr = sqlStr & "  (mallgubun,ReturnCode,ReturnName,ReturnAddress,isusing)"&VbCRLF  '',mappid
		            sqlStr = sqlStr & "  values('lotteCom'"&VbCRLF
		            sqlStr = sqlStr & "  ,'"&html2db(ReturnCode)&"'"&VbCRLF
		            sqlStr = sqlStr & "  ,'"&html2db(ReturnName)&"'"&VbCRLF
		            sqlStr = sqlStr & "  ,'"&html2db(ReturnAddress)&"'"&VbCRLF
		            sqlStr = sqlStr & "  ,'Y'"&VbCRLF
		            '''sqlStr = sqlStr & "  ,'"&html2db(ReturnName)&"'"&VbCRLF
		            sqlStr = sqlStr & "  )"
		            sqlStr = sqlStr & " END"&VbCRLF
		            dbget.Execute sqlStr
		        end if
		        retText = retText & ReturnCode &","&ReturnName&","&ReturnAddress&VbCRLF
		    Next
	        Set ReturnInfo = Nothing
        end if

		''retText = xmlDOM.XML
		Set xmlDOM = Nothing
		
		rw "<script>alert('"&retCnt&" 건 완료');</script>"
	else
		rw "objXML.Status="&objXML.Status
	end if
Set objXML = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->