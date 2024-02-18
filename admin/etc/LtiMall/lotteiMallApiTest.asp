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
	<option value="pumok" <%=CHKIIF(mode="pumok","selected","") %> >��ǰǰ������
	<option value="returnAdr" <%=CHKIIF(mode="returnAdr","selected","") %> >��ǰ�����
	<option value="returnPriceInfo" <%=CHKIIF(mode="returnPriceInfo","selected","") %> >��ۺ���å����
	<option value="brandInfo" <%=CHKIIF(mode="brandInfo","selected","") %> >�귣������
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
    <input type="button" value=" �� �� " onClick="fnSubmit();">
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
    	    'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			'retText = xmlDOM.xml
'rw BinaryToText(objXML.ResponseBody, "euc-kr")

				Dim infoCd, infoDiv, infoDivName, infoItemName, infoDesc, itemMethCd
				Set GoodsArtc = xmlDOM.getElementsByTagName("ArtcItemList")
	        	For each SubNodes in GoodsArtc
	        	    retText= retText & SubNodes.getElementsByTagName("GoodsArtc").item(0).text & "," & SubNodes.getElementsByTagName("GoodsArtcNm").item(0).text
	        	    infoDiv		= Trim(SubNodes.getElementsByTagName("GoodsArtc").item(0).text)						'��ǰ�����ڵ�
	        	    infoDivName = Trim(SubNodes.getElementsByTagName("GoodsArtcNm").item(0).text)					'��ǰ�����̸�
	        	    Set ArtcItem = SubNodes.getElementsByTagName("ArtcItem")
						For each SubSubNodes in ArtcItem
							infoCd			= Trim(SubSubNodes.getElementsByTagName("ItemCd").item(0).text)			'��ǰ�����ڵ�
							infoItemName	= Trim(SubSubNodes.getElementsByTagName("ItemNm").item(0).text)			'��ǰ������
							infoDesc		= Trim(SubSubNodes.getElementsByTagName("ItemDesc").item(0).text)		'��ǰ�󼼼���
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
		objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo & "&ret_no=" & ret_no, false  ''ret_no �ȸ���..
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
		    sqlStr = "update db_temp.dbo.tbl_jaehyumall_returnInfo"
		    sqlStr = sqlStr & " set isusing='N'"
		    sqlStr = sqlStr & " where mallgubun='ltiMall'"
		    sqlStr = sqlStr & " and  isusing='Y'"
		    ''dbget.Execute sqlStr

		    'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			retCnt = xmlDOM.getElementsByTagName("ReturnCount").item(0).text		'�����

			if (retVal) and (retCnt>0) then
                Set ReturnInfo = xmlDOM.getElementsByTagName("ReturnInfo")
    		    for each SubNodes in ReturnInfo
    		        ReturnCode		= Trim(SubNodes.getElementsByTagName("ReturnCode").item(0).text)		'�����/��ǰ���ڵ�
    		        ReturnName      = Trim(SubNodes.getElementsByTagName("ReturnName").item(0).text)		'�����/��ǰ����
    		        ReturnAddress   = Trim(SubNodes.getElementsByTagName("ReturnAddress").item(0).text)		'�����/��ǰ���ּ�
    		        ''if (ReturnCode<>"72125") and (ReturnCode<>"") then ''�⺻�����.
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
			retCnt = xmlDOM.getElementsByTagName("DlvPolcCount").item(0).text		'�����

			If (retVal) and (retCnt>0) Then
                Set ReturnInfo = xmlDOM.getElementsByTagName("DlvPolcInfo")
    		    for each SubNodes in ReturnInfo
    		        DlvPolcNo		= Trim(SubNodes.getElementsByTagName("DlvPolcNo").item(0).text)		'��ۺ���å��ȣ
					EntrNo			= Trim(SubNodes.getElementsByTagName("EntrNo").item(0).text)		'�ŷ�ó��ȣ
					EntrContrNo		= Trim(SubNodes.getElementsByTagName("EntrContrNo").item(0).text)	'�ŷ�ó����ȣ
					LwstEntrNo		= Trim(SubNodes.getElementsByTagName("LwstEntrNo").item(0).text)	'�����ŷ�ó��ȣ
					LwstEntrNm		= Trim(SubNodes.getElementsByTagName("LwstEntrNm").item(0).text)	'�����ŷ�ó��
					Dlex			= Trim(SubNodes.getElementsByTagName("Dlex").item(0).text)			'��ۺ�
					StdAmt			= Trim(SubNodes.getElementsByTagName("StdAmt").item(0).text)		'���رݾ�
					DlexSctCd		= Trim(SubNodes.getElementsByTagName("DlexSctCd").item(0).text)		'��ۺ񱸺��ڵ�
					RtgsDlex		= Trim(SubNodes.getElementsByTagName("RtgsDlex").item(0).text)		'��ǰ��ۺ�
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
		objXML.Open "GET", ltiMallAPIURL & iURI &"?subscriptionId=" & ltiMallAuthNo & "&brnd_nm=�ٹ�����", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			retCnt = xmlDOM.getElementsByTagName("BrandCount").item(0).text		'�����

			If (retVal) and (retCnt>0) Then
                Set ReturnInfo = xmlDOM.getElementsByTagName("BrandInfo")
    		    for each SubNodes in ReturnInfo
    		        BrandCode			= Trim(SubNodes.getElementsByTagName("BrandCode").item(0).text)			'�귣���ڵ�
					BrandName			= Trim(SubNodes.getElementsByTagName("BrandName").item(0).text)			'�귣���
					BrandEnglishName	= Trim(SubNodes.getElementsByTagName("BrandEnglishName").item(0).text)	'�귣�念����
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