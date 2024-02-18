<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim iURI : iURI = request.Form("iURI")
Dim ret_no : ret_no = request.Form("ret_no")
Dim paramname : paramname = request.Form("paramname")
Dim mode : mode = request.Form("mode")
%>
<script language='javascript'>
function fnSubmit(){
    frm = document.frmSubmit;
    frm.submit();
}
function onModeChange(comp){
    var frm = comp.form;
    if (comp.value=="pumok"){
        frm.iURI.value="/openapi/searchGoodsArtcItemCdListOpenApi.lotte";
    }else if (comp.value=="returnAdr"){
        frm.iURI.value="/openapi/searchReturnListOpenApi.lotte";
    }else if (comp.value=="returnPriceInfo"){
        frm.iURI.value="/openapi/searchDlvPolcInfoListOpenApi.lotte";
    }else if (comp.value=="brandInfo"){
        frm.iURI.value="/openapi/searchBrandListOpenApi.lotte";
    }
}
</script>
<table width="800" cellpadding="0" cellspacing="0" border="1" class="a">
<form name="frmSubmit" method="post" action="lotteiMallApiTest.asp">
<select name="mode" onChange="onModeChange(this);">
	<option value="pumok" <%=CHKIIF(mode="pumok","selected","") %> >상품품목정보
	<option value="returnAdr" <%=CHKIIF(mode="returnAdr","selected","") %> >반품배송지
	<option value="returnPriceInfo" <%=CHKIIF(mode="returnPriceInfo","selected","") %> >배송비정책정보
	<option value="brandInfo" <%=CHKIIF(mode="brandInfo","selected","") %> >브랜드정보
</select>
<tr>
    <td>iURI</td>
    <td><input type="text" name="iURI" value="/openapi/searchGoodsArtcItemCdListOpenApi.lotte" size="80" > </td>
</tr>

<tr>
    <td>subscriptionId</td>
    <td>
        <input type="text" name="paramname" value="subscriptionId" size="20" ><br>
        <input type="text" name="subscriptionId" value="<%= ltiMallAuthNo %>" size="80" >
    </td>
</tr>

<tr>
    <td>ret_no</td>
    <td>
        <input type="text" name="paramname" value="ret_no" size="20" ><br>
        <input type="text" name="ret_no" value="<%= ret_no %>" size="80" >
    </td>
</tr>
<tr>
    <td colspan="2" align="center" height="30">
    <input type="button" value=" 서 밋 " onClick="fnSubmit();">
    </td>
</tr>
</form>
</table>
<%
rw iURI
rw ret_no
rw ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo & "&ret_no=" & ret_no

Dim paramnameArr : paramnameArr = split(paramname)
Dim retText, retVal, retCnt
Dim i
Dim ReturnInfo, SubNodes, GoodsArtc, ArtcItem, SubSubNodes
Dim ReturnCode,ReturnName,ReturnAddress
Dim sqlStr

If (mode="pumok") Then
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo , false
    	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    	objXML.Send()

    	If objXML.Status = "200" Then
    	    'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			'retText = xmlDOM.xml
'rw BinaryToText(objXML.ResponseBody, "euc-kr")

				Dim infoCd, infoDiv, infoDivName, infoItemName, infoDesc, itemMethCd
				Set GoodsArtc = xmlDOM.getElementsByTagName("ArtcItemList")
	        	For each SubNodes in GoodsArtc
	        	    retText= retText & SubNodes.getElementsByTagName("GoodsArtc").item(0).text & "," & SubNodes.getElementsByTagName("GoodsArtcNm").item(0).text
	        	    infoDiv		= Trim(SubNodes.getElementsByTagName("GoodsArtc").item(0).text)						'상품구분코드
	        	    infoDivName = Trim(SubNodes.getElementsByTagName("GoodsArtcNm").item(0).text)					'상품구분이름
	        	    Set ArtcItem = SubNodes.getElementsByTagName("ArtcItem")
						For each SubSubNodes in ArtcItem
							infoCd			= Trim(SubSubNodes.getElementsByTagName("ItemCd").item(0).text)			'상품정보코드
							infoItemName	= Trim(SubSubNodes.getElementsByTagName("ItemNm").item(0).text)			'상품정보명
							infoDesc		= Trim(SubSubNodes.getElementsByTagName("ItemDesc").item(0).text)		'상품상세설명
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_jaehyumall_infoCode " & vbcrlf
							sqlStr = sqlStr & " (mallgubun, infoCd, infoDiv, infoDivName, infoItemName, infoDesc, isUsing) " & vbcrlf
       		        	    sqlStr = sqlStr & "  values('ltiMall'" & vbcrlf
        		            sqlStr = sqlStr & " ,'"&html2db(infoCd)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDiv)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDivName)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoItemName)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDesc)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'Y'"&VbCRLF
        		            sqlStr = sqlStr & " )"
'							dbget.Execute sqlStr
							retText= retText + infoCd + "," + infoItemName
						Next
		        	    retText= retText + VbCrlf
				Next
				Set ArtcItem = Nothing
				Set GoodsArtc = Nothing
		    Set xmlDOM = Nothing
		Else
			rw "objXML.Status="&objXML.Status
		End If
    Set objXML = Nothing
ElseIf (mode="returnAdr") Then
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo & "&ret_no=" & ret_no, false  ''ret_no 안먹음..
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
		    sqlStr = "update db_temp.dbo.tbl_jaehyumall_returnInfo"
		    sqlStr = sqlStr & " set isusing='N'"
		    sqlStr = sqlStr & " where mallgubun='ltiMall'"
		    sqlStr = sqlStr & " and  isusing='Y'"
		    ''dbget.Execute sqlStr

		    'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			retCnt = xmlDOM.getElementsByTagName("ReturnCount").item(0).text		'결과수

			if (retVal) and (retCnt>0) then
                Set ReturnInfo = xmlDOM.getElementsByTagName("ReturnInfo")
    		    for each SubNodes in ReturnInfo
    		        ReturnCode		= Trim(SubNodes.getElementsByTagName("ReturnCode").item(0).text)		'출고지/반품지코드
    		        ReturnName      = Trim(SubNodes.getElementsByTagName("ReturnName").item(0).text)		'출고지/반품지명
    		        ReturnAddress   = Trim(SubNodes.getElementsByTagName("ReturnAddress").item(0).text)		'출고지/반품지주소
    		        ''if (ReturnCode<>"72125") and (ReturnCode<>"") then ''기본배송지.
    		        if (ReturnCode<>"") then 
    		            sqlStr = "if Exists(select * from db_temp.dbo.tbl_jaehyumall_returnInfo where mallgubun='ltiMall' and ReturnCode='"&ReturnCode&"') "&VbCRLF
    		            sqlStr = sqlStr & " BEGIN"&VbCRLF
    		            sqlStr = sqlStr & "  update db_temp.dbo.tbl_jaehyumall_returnInfo"&VbCRLF
    		            sqlStr = sqlStr & "  set ReturnName='"&html2db(ReturnName)&"'"&VbCRLF
    		            sqlStr = sqlStr & "  , ReturnAddress='"&html2db(ReturnAddress)&"'"&VbCRLF
    		            sqlStr = sqlStr & "  , isUsing='Y'"&VbCRLF
    		            sqlStr = sqlStr & "  , lastupdate=getdate()"&VbCRLF
    		            sqlStr = sqlStr & "  where mallgubun='ltiMall'"&VbCRLF
    		            sqlStr = sqlStr & "  and ReturnCode='"&ReturnCode&"'"&VbCRLF
    		            sqlStr = sqlStr & " END"&VbCRLF
    		            sqlStr = sqlStr & " ELSE "&VbCRLF
    		            sqlStr = sqlStr & " BEGIN"&VbCRLF
    		            sqlStr = sqlStr & "  Insert into db_temp.dbo.tbl_jaehyumall_returnInfo"&VbCRLF
    		            sqlStr = sqlStr & "  (mallgubun,ReturnCode,ReturnName,ReturnAddress,isusing)"&VbCRLF
    		            sqlStr = sqlStr & "  values('ltiMall'"&VbCRLF
    		            sqlStr = sqlStr & "  ,'"&html2db(ReturnCode)&"'"&VbCRLF
    		            sqlStr = sqlStr & "  ,'"&html2db(ReturnName)&"'"&VbCRLF
    		            sqlStr = sqlStr & "  ,'"&html2db(ReturnAddress)&"'"&VbCRLF
    		            sqlStr = sqlStr & "  ,'Y'"&VbCRLF
    		            sqlStr = sqlStr & "  )"
    		            sqlStr = sqlStr & " END"&VbCRLF
    		            ''dbget.Execute sqlStr
    		        end if
    		        retText = retText & ReturnCode &","&ReturnName&","&ReturnAddress&VbCRLF
    		    Next
		        Set ReturnInfo = Nothing
            end if

			''retText = xmlDOM.XML
			Set xmlDOM = Nothing
		else
			rw "objXML.Status="&objXML.Status
		end if
    Set objXML = Nothing
ElseIf (mode="returnPriceInfo") Then
Dim DlvPolcNo, EntrNo, EntrContrNo, LwstEntrNo, LwstEntrNm, Dlex, StdAmt, DlexSctCd, RtgsDlex
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			retCnt = xmlDOM.getElementsByTagName("DlvPolcCount").item(0).text		'결과수

			If (retVal) and (retCnt>0) Then
                Set ReturnInfo = xmlDOM.getElementsByTagName("DlvPolcInfo")
    		    for each SubNodes in ReturnInfo
    		        DlvPolcNo		= Trim(SubNodes.getElementsByTagName("DlvPolcNo").item(0).text)		'배송비정책번호
					EntrNo			= Trim(SubNodes.getElementsByTagName("EntrNo").item(0).text)		'거래처번호
					EntrContrNo		= Trim(SubNodes.getElementsByTagName("EntrContrNo").item(0).text)	'거래처계약번호
					LwstEntrNo		= Trim(SubNodes.getElementsByTagName("LwstEntrNo").item(0).text)	'하위거래처번호
					LwstEntrNm		= Trim(SubNodes.getElementsByTagName("LwstEntrNm").item(0).text)	'하위거래처명
					Dlex			= Trim(SubNodes.getElementsByTagName("Dlex").item(0).text)			'배송비
					StdAmt			= Trim(SubNodes.getElementsByTagName("StdAmt").item(0).text)		'기준금액
					DlexSctCd		= Trim(SubNodes.getElementsByTagName("DlexSctCd").item(0).text)		'배송비구분코드
					RtgsDlex		= Trim(SubNodes.getElementsByTagName("RtgsDlex").item(0).text)		'반품배송비
    		        retText = retText & DlvPolcNo &","&EntrNo&","&EntrContrNo&","&LwstEntrNo&","&LwstEntrNm&","&Dlex&","&StdAmt&","&DlexSctCd&","&RtgsDlex&VBCRLF
    		    Next
		        Set ReturnInfo = Nothing
            end if
			Set xmlDOM = Nothing
		else
			rw "objXML.Status="&objXML.Status
		end if
    Set objXML = Nothing
ElseIf (mode="brandInfo") Then
Dim BrandCode, BrandName, BrandEnglishName
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo & "&brnd_nm=텐바이텐", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			retCnt = xmlDOM.getElementsByTagName("BrandCount").item(0).text		'결과수

			If (retVal) and (retCnt>0) Then
                Set ReturnInfo = xmlDOM.getElementsByTagName("BrandInfo")
    		    for each SubNodes in ReturnInfo
    		        BrandCode			= Trim(SubNodes.getElementsByTagName("BrandCode").item(0).text)			'브랜드코드
					BrandName			= Trim(SubNodes.getElementsByTagName("BrandName").item(0).text)			'브랜드명
					BrandEnglishName	= Trim(SubNodes.getElementsByTagName("BrandEnglishName").item(0).text)	'브랜드영문명
    		        retText = retText & BrandCode &","&BrandName&","&BrandEnglishName&VBCRLF
    		    Next
		        Set ReturnInfo = Nothing
            end if
			Set xmlDOM = Nothing
		else
			rw "objXML.Status="&objXML.Status
		end if
    Set objXML = Nothing
End If
%>

<table width="800" cellpadding="0" cellspacing="0" border="1" class="a">
<tr>
    <td><textarea cols="100" rows="80"><%= retText %></textarea></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->