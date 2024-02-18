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
<option value="pumok" <%=CHKIIF(mode="pumok","selected","") %> >��ǰǰ������
<option value="returnAdr" <%=CHKIIF(mode="returnAdr","selected","") %> >��ǰ�����
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
    <input type="button" value=" �� �� " onClick="fnSubmit();">
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
    	    'XML�� ���� DOM ��ü ����
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
			retVal = xmlDOM.LoadXML(BinaryToText(objXML.ResponseBody, "euc-kr"))
			''retText = xmlDOM.xml

				Dim infoCd, infoDiv, infoDivName, infoItemName, infoDesc, itemMethCd
				Set GoodsArtc = xmlDOM.getElementsByTagName("GoodsArtc")
	        	For each SubNodes in GoodsArtc
	        	    retText= retText + SubNodes.getAttribute("code") + "," + SubNodes.getAttribute("code_nm")
	        	    infoDiv		= Trim(SubNodes.getAttribute("code"))						'��ǰ�����ڵ�
	        	    infoDivName = Trim(SubNodes.getAttribute("code_nm"))					'��ǰ�����̸�

	        	    Set ArtcItem = SubNodes.getElementsByTagName("ArtcItem")
						For each SubSubNodes in ArtcItem
							
							infoCd			= Trim(SubSubNodes.getAttribute("itemCd"))			'��ǰ�����ڵ�
							infoItemName	= Trim(SubSubNodes.getAttribute("itemNm"))			'��ǰ������
							infoDesc		= Trim(SubSubNodes.getAttribute("itemDesc"))		'��ǰ�󼼼���
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
        ''<![CDATA[410570 ��� ���� �ϻ굿�� ������ 418-11����]]>
        ''</ReturnAddress>
        ''</ReturnInfo>

    		objXML.Open "GET", lotteAPIURL & iURI &"?subscriptionId=" & lotteAuthNo & "&ret_no=" & ret_no, false  ''ret_no �ȸ���..
    		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    		objXML.Send()
    		If objXML.Status = "200" Then
    		    sqlStr = "update db_temp.dbo.tbl_jaehyumall_returnInfo"
    		    sqlStr = sqlStr & " set isusing='N'"
    		    sqlStr = sqlStr & " where mallgubun='lotteCom'"
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
        		        ReturnCode		= Trim(SubNodes.getElementsByTagName("ReturnCode").item(0).text)		'������Ϸù�ȣ
        		        ReturnName      = Trim(SubNodes.getElementsByTagName("ReturnName").item(0).text)		'�������
        		        ReturnAddress   = Trim(SubNodes.getElementsByTagName("ReturnAddress").item(0).text)		'�����ȣ,�⺻�ּ�,���ּ�
        		        '�����
        		        '����ó
        		        '�޴���
        		        ''if (ReturnCode<>"72125") and (ReturnCode<>"") then ''�⺻�����.
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