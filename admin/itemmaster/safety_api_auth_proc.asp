<%@ language=vbscript %>
<% option explicit %>
<% response.charset = "euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
	Dim objHttp, sData, url, vCertNum, vGetData, vLastMessage, vIsSave, vExistMSG, vItemID
	dim oResult, resultCode, resultMsg, oDetailData, vQuery, vIdx, x, safetydiv
	dim certUid, certOrganName, certState, certDiv, certDate, certChgDate, certChgReason, firstCertNum, productName, brandName
	dim modelName, categoryName, importDiv, makerName, makerCntryName, importerName, certificationImageUrls, statusmode
	safetydiv = requestCheckVar(request("safetydiv"),300)
	vCertNum = requestCheckVar(request("certnum"),300)
	vIsSave = requestCheckVar(request("issave"),1)
	vItemID = requestCheckVar(request("itemid"),20)
	statusmode = requestCheckVar(request("statusmode"),16)
	'vCertNum = "SU071356-12001"	''����
	'vCertNum = "SU071323-14001"	''����
	'vCertNum = "SU071677-14001"	''���
	'vCertNum = "JU07653-12001A"	''����
	'vCertNum = "JH07282-6002"	''Data not found

	If vIsSave = "u" Then
		vCertNum = fnRealCertNumSetting(vItemID, vCertNum, statusmode)
		vIsSave = "o"
	End If

	If vIsSave = "x" Then	'### �ܼ� ��ȸ.
		
		'vExistMSG = fnCheckCertNum(vCertNum)
		vExistMSG = ""	' �ߺ�üũ ����.
		
		If vExistMSG <> "" Then
			vLastMessage = vExistMSG
		Else
			vGetData = "certNum={" & vCertNum & "}"
			url = "http://www.safetykorea.kr/openapi/api/cert/certificationDetail.json?certNum="&vCertNum&""

			Set objHttp = server.CreateObject("MSXML2.ServerXMLHTTP")
			If IsNull(objHttp) Then
				vLastMessage = "���� ���� ����"
			Else
				
				objHttp.Open "Get", url, False
				objHttp.SetRequestHeader "Authkey", "46aeb476-f79d-423f-95ea-109feeb0ee91"
				objHttp.Send
				
				If objHttp.Status = 200 Then
				    sData = objHttp.responseText

					If sData <> "" Then
						set oResult = JSON.parse(sData)
						resultCode = oResult.resultCode
						resultMsg = oResult.resultMsg
						
						If resultCode = "2000" Then
							Set oDetailData = oResult.resultData
							certState = oDetailData.certState
							Set oDetailData = Nothing

							vLastMessage = certState
							
						ElseIf resultCode = "2004" Then
							vLastMessage = "����� �ȵ� ������ȣ�Դϴ�."	'No Data
						ElseIf resultCode = "4000" Then
							vLastMessage = "�Ұ����� ����Ű�Դϴ�."	'Invalid Auth Key
						ElseIf resultCode = "4001" Then
							vLastMessage = "�Ұ����� IP�Դϴ�."	'Invalid IP
						ElseIf resultCode = "4005" Then
							vLastMessage = "�Ұ����� �Ķ���Ͱ��Դϴ�."	'Invalid Parameter
						ElseIf resultCode = "5000" Then
							vLastMessage = "SAFETY KOREA ���� �����Դϴ�."	'Internal Server Error. Message:{}
						Else
							vLastMessage = "Error 1"
						End If
						
						set oResult = Nothing
					Else
						vLastMessage = "Error 2"
					End If
				Else
				    vLastMessage = "Status: " & objHttp.Status & " | " & objHttp.responseText
				End If
			End If
			
			Set objHttp = Nothing
		End If
		
	ElseIf vIsSave = "o" Then	'### ��ǰ��������ư Ŭ���� API�� �ٽ� ���� ������ �����ϰ� ������ idx�� �Ѱ���(���δ�������� ������ �̰�����).
		
		Dim i, vTmpCert, vTmpIdx
		vTmpCert = Split(vCertNum,",")
		
		For i = LBound(vTmpCert) To UBound(vTmpCert)
			'### x ���� ������ ���ռ� Ȯ�� ���� ������ȣ���� ����. x�� ���� ������ ����.
			'### �׷��� tbl_safetycert_tenReg �� ���̺��� �����ϰ�, tbl_safetycert_info �̰����� �������� ����.
			
			If vTmpCert(i) <> "x" Then
				vGetData = "certNum={" & vTmpCert(i) & "}"
				url = "http://www.safetykorea.kr/openapi/api/cert/certificationDetail.json?certNum="&vTmpCert(i)&""

				Set objHttp = server.CreateObject("MSXML2.ServerXMLHTTP")
				If IsNull(objHttp) Then
					vLastMessage = "���� ���� ����"
				Else
					
					objHttp.Open "Get", url, False
					objHttp.SetRequestHeader "Authkey", "46aeb476-f79d-423f-95ea-109feeb0ee91"
					objHttp.Send
					
					If objHttp.Status = 200 Then
					    sData = objHttp.responseText
					    'response.write strJsonText
					    
						If sData <> "" Then
							set oResult = JSON.parse(sData)
							resultCode = oResult.resultCode
							resultMsg = oResult.resultMsg
							
							If resultCode = "2000" Then
								Set oDetailData = oResult.resultData
								certUid = oDetailData.certUid
								certOrganName = oDetailData.certOrganName
								certState = oDetailData.certState
								certDiv = oDetailData.certDiv
								certDate = oDetailData.certDate
								certChgDate = oDetailData.certChgDate
								certChgReason = oDetailData.certChgReason
								firstCertNum = oDetailData.firstCertNum
								productName = oDetailData.productName
								brandName = oDetailData.brandName
								modelName = oDetailData.modelName
								categoryName = oDetailData.categoryName
								importDiv = oDetailData.importDiv
								makerName = oDetailData.makerName
								makerCntryName = oDetailData.makerCntryName
								importerName = oDetailData.importerName
								certificationImageUrls = oDetailData.certificationImageUrls
								Set oDetailData = Nothing
								
								vQuery = "INSERT INTO db_temp.[dbo].[tbl_safetycert_info_temp](certUid,certOrganName,certNum,certState,certDiv,certDate,certChgDate,"
								vQuery = vQuery & "certChgReason,firstCertNum,productName,brandName,modelName,categoryName,importDiv,makerName,makerCntryName,importerName) "
								vQuery = vQuery & "VALUES('" & certUid & "','" & certOrganName & "','" & trim(vTmpCert(i)) & "','" & certState & "','" & certDiv & "','" & certDate & "',"
								vQuery = vQuery & "'" & certChgDate & "','" & certChgReason & "','" & firstCertNum & "','" & productName & "','" & brandName & "',"
								vQuery = vQuery & "'" & modelName & "','" & categoryName & "','" & importDiv & "','" & makerName & "','" & makerCntryName & "','" & importerName & "')"
								dbget.execute(vQuery)

								vQuery = " SELECT SCOPE_IDENTITY() "
								rsget.CursorLocation = adUseClient
								rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
						 		IF Not rsget.EOF THEN
						 			vIdx = rsget(0)
						 		END IF
						 		rsget.close
								
								If certificationImageUrls <> "" Then
									vQuery = ""
									For x = LBound(Split(certificationImageUrls,",")) To UBound(Split(certificationImageUrls,","))
										vQuery = vQuery & "INSERT INTO db_temp.[dbo].[tbl_safetycert_image_temp](topidx,certNum,ImageUrls) "
										vQuery = vQuery & "VALUES('" & vIdx & "','" & trim(vTmpCert(i)) & "','" & Split(certificationImageUrls,",")(x) & "'); "
									Next
									
									If vQuery <> "" Then
										dbget.execute(vQuery)
									End If
								End If

								vTmpIdx = vTmpIdx & vIdx & ","
								
							End If
							
							set oResult = Nothing
						End If
					End If
				End If
				
				Set objHttp = Nothing
			End If
		Next
		
		If vTmpIdx <> "" Then
			If Right(vTmpIdx,1) = "," Then
				vTmpIdx = Left(vTmpIdx, Len(vTmpIdx)-1)
			End If
			vLastMessage = vTmpIdx
		End If
	End If
	
	Response.Write vLastMessage
	
'### �̹� ����� �� ������ȣ���� üũ �Լ�.
Function fnCheckCertNum(cn)
	Dim sql, r
	sql = "SELECT count(itemid) From db_item.[dbo].[tbl_safetycert_info] where certNum = '" & cn & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
	if rsget(0) > 0 then
		r = "�Էµ� ������ȣ�� ���� MD����ó�� �Ϸ�� ��ǰ�߿� �̹� ��ϵ� ������ȣ�Դϴ�."
	end if
	rsget.close
	
	If r = "" Then
		sql = "SELECT count(itemid) From db_temp.[dbo].[tbl_safetycert_info_waititem] where certNum = '" & cn & "'"
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
		if rsget(0) > 0 then
			r = "�Էµ� ������ȣ�� ���� MD����ó�� ���ܰ� ��ǰ�߿� �̹� ��ϵ� ������ȣ�Դϴ�."
		end if
		rsget.close
	End If
	
	fnCheckCertNum = r
End Function

Function fnRealCertNumSetting(itemid, cn, statusmode)
	Dim sql, r, i, vCompare, vString
	
	vString = Split(cn,",")

	if statusmode="wait" then
		sql = "select certNum from db_temp.[dbo].[tbl_safetycert_info_waititem] where itemid = '" & itemid & "'"
	else
		sql = "select certNum from db_temp.[dbo].[tbl_safetycert_info] where itemid = '" & itemid & "'"
	end if

	'response.write sql & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
	if not rsget.eof then
		do until rsget.eof
			if vCompare = "" Then
				vCompare = vCompare & rsget(0)
			else
				vCompare = vCompare & "," & rsget(0)
			end if
			
			rsget.movenext
		loop
	end if
	rsget.close
	
	For i = LBound(vString) To UBound(vString)
		if instr(vCompare, vString(i)) < 1 then
			r = r & vString(i) & ","
		end if
	Next
	
	if r <> "" then
		if right(r,1) = "," then
			r = left(r, len(r)-1)
		end if
	end if
	
	fnRealCertNumSetting = r

End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->