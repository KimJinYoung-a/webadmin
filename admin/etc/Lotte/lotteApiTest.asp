<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
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
    }
}
</script>
<table width="800" cellpadding="0" cellspacing="0" border="1" class="a">
<form name="frmSubmit" method="post" action="lotteApiTest.asp">
<select name="mode" onChange="onModeChange(this);">
<option value="pumok" <%=CHKIIF(mode="pumok","selected","") %> >상품품목정보
<option value="returnAdr" <%=CHKIIF(mode="returnAdr","selected","") %> >반품배송지
</select>
<tr>
    <td>iURI</td>
    <td><input type="text" name="iURI" value="/openapi/searchGoodsArtcItemCdListOpenApi.lotte" size="80" > </td>
</tr>

<tr>
    <td>subscriptionId</td>
    <td>
        <input type="text" name="paramname" value="subscriptionId" size="20" ><br>
        <input type="text" name="subscriptionId" value="<%= lotteAuthNo %>" size="80" >
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
rw lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo & "&ret_no=" & ret_no

Dim paramnameArr : paramnameArr = split(paramname)
Dim retText, retVal, retCnt
Dim i
Dim ReturnInfo, SubNodes, GoodsArtc, ArtcItem, SubSubNodes
Dim ReturnCode,ReturnName,ReturnAddress
Dim sqlStr

If (mode="pumok") Then
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objXML.Open "GET", lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo , false
    	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    	objXML.Send()

    	If objXML.Status = "200" Then
    	    'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			''retText = xmlDOM.xml

				Dim infoCd, infoDiv, infoDivName, infoItemName, infoDesc, itemMethCd
				Set GoodsArtc = xmlDOM.getElementsByTagName("GoodsArtc")
	        	For each SubNodes in GoodsArtc
	        	    retText= retText + SubNodes.getAttribute("code") + "," + SubNodes.getAttribute("code_nm")
	        	    infoDiv		= Trim(SubNodes.getAttribute("code"))						'상품구분코드
	        	    infoDivName = Trim(SubNodes.getAttribute("code_nm"))					'상품구분이름

	        	    Set ArtcItem = SubNodes.getElementsByTagName("ArtcItem")
						For each SubSubNodes in ArtcItem
							
							infoCd			= Trim(SubSubNodes.getAttribute("itemCd"))			'상품정보코드
							infoItemName	= Trim(SubSubNodes.getAttribute("itemNm"))			'상품정보명
							infoDesc		= Trim(SubSubNodes.getAttribute("itemDesc"))		'상품상세설명
                            itemMethCd      = Trim(SubSubNodes.getAttribute("itemMethCd"))
                            
							sqlStr = "if Exists(select * from db_temp.dbo.tbl_jaehyumall_infoCode where mallgubun='lotteCom' and infoCd='"&infoCd&"') "&VbCRLF
        		            sqlStr = sqlStr & " BEGIN"&VbCRLF
        		            sqlStr = sqlStr & "	UPDATE db_temp.dbo.tbl_jaehyumall_infoCode "&VbCRLF
        		            sqlStr = sqlStr & "	SET infoDivName='"&html2db(infoDivName)&"'"&VbCRLF
        		            sqlStr = sqlStr & "	, infoItemName='"&html2db(infoItemName)&"'"&VbCRLF
        		            sqlStr = sqlStr & "	, infoDesc='"&html2db(infoDesc)&"'"&VbCRLF
        		            sqlStr = sqlStr & "	, isUsing='Y'"&VbCRLF
        		            sqlStr = sqlStr & "	where mallgubun='lotteCom'"&VbCRLF
        		            sqlStr = sqlStr & " and infoCd='"&infoCd&"'"&VbCRLF
        		            sqlStr = sqlStr & " END"&VbCRLF
        		            sqlStr = sqlStr & " ELSE "&VbCRLF
        		            sqlStr = sqlStr & " BEGIN"&VbCRLF
							sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_jaehyumall_infoCode " & vbcrlf
							sqlStr = sqlStr & " (mallgubun, infoCd, infoDiv, infoDivName, infoItemName, infoDesc, isUsing) " & vbcrlf
       		        	    sqlStr = sqlStr & "  values('lotteCom'" & vbcrlf
        		            sqlStr = sqlStr & " ,'"&html2db(infoCd)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDiv)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDivName)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoItemName)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'"&html2db(infoDesc)&"'" & VbCRLF
        		            sqlStr = sqlStr & " ,'Y'"&VbCRLF
        		            sqlStr = sqlStr & " )"
							sqlStr = sqlStr & " END"&VbCRLF
							''dbget.Execute sqlStr
							retText= retText + SubSubNodes.getAttribute("itemCd") + "," + SubSubNodes.getAttribute("itemNm") + "{"+itemMethCd+"}"
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
        ''<ReturnInfo>
        ''<ReturnCode>115220</ReturnCode>
        ''<ReturnName>
        ''<![CDATA[ ETHOS]]>
        ''</ReturnName>
        ''<ReturnAddress>
        ''<![CDATA[410570 경기 고양시 일산동구 성석동 418-11번지]]>
        ''</ReturnAddress>
        ''</ReturnInfo>

    		objXML.Open "GET", lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo & "&ret_no=" & ret_no, false  ''ret_no 안먹음..
    		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    		objXML.Send()
    		If objXML.Status = "200" Then
    		    sqlStr = "update db_temp.dbo.tbl_jaehyumall_returnInfo"
    		    sqlStr = sqlStr & " set isusing='N'"
    		    sqlStr = sqlStr & " where mallgubun='lotteCom'"
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
        		        ReturnCode		= Trim(SubNodes.getElementsByTagName("ReturnCode").item(0).text)		'배송지일련번호
        		        ReturnName      = Trim(SubNodes.getElementsByTagName("ReturnName").item(0).text)		'배송지명
        		        ReturnAddress   = Trim(SubNodes.getElementsByTagName("ReturnAddress").item(0).text)		'우편번호,기본주소,상세주소
        		        '담당자
        		        '연락처
        		        '휴대폰
        		        ''if (ReturnCode<>"72125") and (ReturnCode<>"") then ''기본배송지.
        		        if (ReturnCode<>"") then 
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
        		            sqlStr = sqlStr & "  (mallgubun,ReturnCode,ReturnName,ReturnAddress,isusing)"&VbCRLF
        		            sqlStr = sqlStr & "  values('lotteCom'"&VbCRLF
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
End If
%>

<table width="800" cellpadding="0" cellspacing="0" border="1" class="a">
<tr>
    <td><textarea cols="100" rows="80"><%= retText %></textarea></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->