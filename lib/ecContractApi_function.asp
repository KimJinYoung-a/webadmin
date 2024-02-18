
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

'###########################################################
' Description : ���ڰ�� ���� api
' Hieditor : 2017.07.05 ������
'###########################################################

'########################### ���÷����� api ####################################################  
public  FecURL, FecID,FecAUser
public Fec_defctrtype , Fec_defctrtype_M,Fec_addctrtype,Fec_addctrtype_M

 if (application("Svr_Info")	= "Dev") then
	FecURL= "http://webtax-test.uplus.co.kr" 
	FecID = "CTENTEN"
	FecAUser = "CTENTEN"
	Fec_defctrtype ="429"
	Fec_defctrtype_M ="432"
	Fec_addctrtype = "435"
	Fec_addctrtype_M = "436"
else
	FecURL = "https://edocu.uplus.co.kr" 
	FecID = "CUBE1010"
	FecAUser = "CTENTEN"
	Fec_defctrtype ="6298"
	Fec_defctrtype_M ="2269"
	Fec_addctrtype = "2267"
	Fec_addctrtype_M = "2268"
end if
 
'########################### Docu�� api ####################################################   
 public  FecDocuURL, adminApiKey
 if (application("Svr_Info")	= "Dev") then
	FecDocuURL= "https://testadminapi.10x10.co.kr" 
else
	FecDocuURL = "https://adminapi.10x10.co.kr:31443"
	adminApiKey = "71ec2fbe40cfbcc43882e1141f662fk9e291ede5bdf9ed6a4defc28"
end if


 public Faccess_token 
 public Ftoken_type
 public Frefresh_token
 public Fchkerror
 public FErrMsg
 
'########################### ��ū ���� #################################################### 
 	public Sub sbGetNewToken(ecId,ecPwd)
 	dim APIpath,strParam
 	dim objXML, iRbody,jsResult
 	dim sqlStr
 	 
			APIpath = FecURL&"/oauth/token"
			strParam = "?grant_type=password&client_id=edocuopenapi&username="&ecId&"&password="&ecPwd
		 	'On Error Resume Next
				Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
				    objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "text/html"
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
					 response.write iRbody
					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							Faccess_token	= jsResult.access_token
							Ftoken_type		= jsResult.token_type
							Frefresh_token	= jsResult.refresh_token						
						Set jsResult = Nothing
					End If
				Set objXML = Nothing
		 	'On Error Goto 0 
		 	
		 	if Faccess_token <> "" Then
				sqlStr = "insert into db_partner.dbo. tbl_partner_ctrLg_token(access_token, refresh_token)"	
				sqlStr = sqlStr & "values('"&Faccess_token&"','"&Frefresh_token&"')"
				dbget.Execute sqlStr
			end if	
	End Sub		

 
'########################### ��ū ���� #################################################### 
 	public Sub sbGetRefToken(reftoken)
 	dim APIpath,strParam
 	dim objXML, iRbody,jsResult
 	dim sqlStr
			APIpath = FecURL&"/oauth/token"
			strParam = "?grant_type=refresh_token&client_id=edocuopenapi&refresh_token="&reftoken
		 	On Error Resume Next
				Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
				    objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "text/html"
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
					 
					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							Faccess_token	= jsResult.access_token
							Ftoken_type		= jsResult.token_type
							Frefresh_token	= jsResult.refresh_token						
						Set jsResult = Nothing
					End If
				Set objXML = Nothing
		 	On Error Goto 0 
		 	
		 	if Faccess_token <> "" Then
				sqlStr = "insert into db_partner.dbo. tbl_partner_ctrLg_token(access_token, refresh_token)"	
				sqlStr = sqlStr & "values('"&Faccess_token&"','"&Frefresh_token&"')"
				dbget.Execute sqlStr
			end if	
	End Sub		
	
'########################### ȸ��Ȯ�� #################################################### 	
	public Function  fnCheckUser(bcompno,access_token)
	 	dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		dim userStatus 
				APIpath = FecURL&"/api/checkuser"
				strParam = "?corp_id="&bcompno
					 
				On Error Resume Next
					Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
					    objXML.Open "GET", APIpath&strParam , False
						objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
						objXML.SetRequestHeader "Authorization", "Bearer " & access_token
						objXML.Send()
						iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")				
						iRbody= replace(iRbody,"tmpCallBack({","{")
						iRbody = replace(iRbody,"})","}") 
						If objXML.Status = "200" Then 
							Set jsResult = JSON.parse(iRbody)
								userStatus	= jsResult.status_cd
								Fchkerror      = jsResult.error
							Set jsResult = Nothing			
						end if				
					Set objXML = Nothing
				On Error Goto 0
				
				fnCheckUser = userStatus
	End Function 
	
	
	'########################### ������  #################################################### 	
	public Function fnCheckCont(ecCtrSeq,B_COMPANY_NO,ecBUser, access_token)
		dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		
		APIpath =FecURL&"/api/checkCont" 
		strParam = "?seq="&ecCtrSeq&"&venderno="&B_COMPANY_NO&"&users="&server.URLEncode(ecBUser)
		
			On Error Resume Next
		
		Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
			  objXML.Open "GET", APIpath&strParam , False
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.SetRequestHeader "Authorization", "Bearer " & access_token
				objXML.Send()				
				iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			
				iRbody= replace(iRbody,"tmpCallBack({","{")
			 	iRbody = replace(iRbody,"})","}") 
		
				If objXML.Status = "200" Then
					Set jsResult = JSON.parse(iRbody)
						con_status	= jsResult.status
						con_info= jsResult.info 
						Fchkerror = jsResult.error
						if con_status ="fail" Then 
								 if (con_info="001") then
								 		FErrMsg= "��༭��ȣ �� ����"
								 elseif (con_info="002")then
								  	FErrMsg= "����ڹ�ȣ(venderno) �� ����,venderno �� ����� �������� ����"
								 elseif (con_info="003")then
								  	FErrMsg= "�ش� ������ ���� ������ ������������" 
								 elseif (con_info="004")then
									  FErrMsg= "�ش繮���� �̹� �����û ����" 														  
								 end if 
								  
						end if
						
				Set jsResult = Nothing
				End If
			Set objXML = Nothing
			
				On Error Goto 0 
				fnCheckCont = con_status
	end function		
	
	'########################### ����Ȯ��  #################################################### 	
 	public Function fnViewEcCont(ecCtrseq, bcompno, ecBuser,access_token)			 
 	dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		dim con_status,con_info,con_error,ecCtrState
	APIpath =FecURL&"/api/viewCont" 
	strParam = "?seq="&ecCtrseq&"&venderno="&bcompno&"&users=" &server.URLEncode(ecBUser)
	
 	On Error Resume Next 
 Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
		  objXML.Open "GET", APIpath&strParam , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.SetRequestHeader "Authorization", "Bearer " & access_token
			objXML.Send()				
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
		
			iRbody= replace(iRbody,"tmpCallBack({","{")
		 	iRbody = replace(iRbody,"})","}") 
  
			If objXML.Status = "200" Then
				Set jsResult = JSON.parse(iRbody)
					con_status	= jsResult.status
					con_info= jsResult.info 
				'	con_error = jsResult.error
					if con_status ="fail" Then 
							 if (con_info="001") then
							 		FErrMsg= "��༭��ȣ �� ����"
							 elseif (con_info="002")then
							  	FErrMsg= "����ڹ�ȣ(venderno) �� ����,venderno �� ����� �������� ����"
							 elseif (con_info="003")then
							  	FErrMsg= "�ش� ������ ���� ������ ������������"  
							 elseif (con_info="004")then
							  	FErrMsg= "�ش繮���� �����Ǿ����ϴ�."   	
							  	ecCtrState = "-1"
							 end if 
					else
						ecCtrState = jsResult.nowstat_vl			  
					end if
					
			Set jsResult = Nothing
			End If
		Set objXML = Nothing		 
	On Error Goto 0 
	fnViewEcCont = ecCtrState
	
	End Function

	'########################### ����Ȯ��  #################################################### 	
 	Function fnViewEcContInfo(ByVal ecCtrseq, ByVal bcompno, ByVal ecBuser, ByVal access_token, ByRef FErrMsg)			 
		dim APIpath, strParam
		dim objXML, iRbody,jsResult
		dim con_status, con_info, con_error, ecCtrState
		APIpath =FecURL&"/api/viewCont" 
		strParam = "?seq="&ecCtrseq&"&venderno="&bcompno&"&users=" &server.URLEncode(ecBUser)
		
		On Error Resume Next 
		Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
			objXML.Open "GET", APIpath&strParam , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.SetRequestHeader "Authorization", "Bearer " & access_token
			objXML.Send()				
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			
			iRbody= replace(iRbody,"tmpCallBack({","{")
			iRbody = replace(iRbody,"})","}") 
	
			If objXML.Status = "200" Then
				Set jsResult = JSON.parse(iRbody)
				con_status	= jsResult.status
				con_info= jsResult.info 
			'	con_error = jsResult.error
				if con_status ="fail" Then 
					if (con_info="001") then
						FErrMsg= "��༭��ȣ �� ����"
					elseif (con_info="002")then
						FErrMsg= "����ڹ�ȣ(venderno) �� ����,venderno �� ����� �������� ����"
					elseif (con_info="003")then
						FErrMsg= "�ش� ������ ���� ������ ������������"  
					elseif (con_info="004")then
						FErrMsg= "�ش繮���� �����Ǿ����ϴ�."   	
						ecCtrState = "-1"
					end if 
				else
					ecCtrState = jsResult.nowstat_vl			  
				end if
				Set jsResult = Nothing
			End If
		Set objXML = Nothing		 
		On Error Goto 0 
		fnViewEcContInfo = ecCtrState
	End Function
%>

	 