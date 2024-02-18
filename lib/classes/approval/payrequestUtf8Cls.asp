<%
Class CPayRequest
public FreportIdx
public FadminId
public FpayrequestIdx
public Farap_cd			
public FreportName      
public FreportPrice     
public Fscmlinkno       
public Fbigo            
public Freportcontents  
public Freportstate     
public Freferid         
 
public Fregdate         
public Fedmsidx
public Farap_nm     
public Facc_cd    
public Facc_use_cd  
public Facc_nm          
public FedmsName        
public Fedmscode        
public FlastApprovalid   
                 
public Fpayrequestdate  
public Fpayrequestprice 
public FinBank          
public FaccountNo       
public FaccountHolder   
public Fpaydate         
public FoutBank         
public Fpayrealdate     
public Fpayrealprice    
public Fyyyymm          
public FisTakeDoc       
public Fpayrequeststate  
public FpayComment

public FsumPayRequestPRice
public Fsumpayrealprice

public FisLast
public Fauthstate
public Fauthposition

public Fusername
public Fpartname

public FPageSize
public FCurrPage
public FSPageNo
public FEPageNo
public FTotCnt
  
public FpayRequestTitle
public FoutBankName

public FpayDocIdx  
public Fpaydockind 
public Fvatkind  	
public Fissuedate  
public Fitemname  	
public Ftotprice 	
public Fsupplyprice
public Fvatprice  	
public Fetaxkey  	
public FDocbigo  			
public Fattachfile 

public Fcust_cd
public Fcust_nm  
public FBiz_no

public Fpaytype
public Fcurrencytype
public Fcurrencyprice

public FerpLinkType
public FACC_GRP_CD

public FerpDocLinkType
public FerpDocLinkKey

	'//결제요청서 리스트 가져오기
	public Function fnGetPayRequestList
	Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getListCnt]('"&FadminId&"',"&Fpayrequeststate&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayRequest_getList('"&FadminId&"',"&Fpayrequeststate&","&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayRequestList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
		'//결제요청서 세금계산서 차후수취 처리된 리스트 가져오기
	public Function fnGetPayRequestDocList
	Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayDoc_getListCnt]('"&FadminId&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayDoc_getList('"&FadminId&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayRequestDocList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	
	'//결제요청서 기본정보 내용보기
	public Function fnGetPayRequestData
	Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getData]( "&FreportIdx&", "&FpayrequestIdx&")"   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			Farap_cd			=rsget("arap_cd")
			FreportName         =rsget("reportName")    
			FreportPrice        =rsget("reportPrice")   
			Fscmlinkno          =rsget("scmlinkno")     
			Fbigo               =rsget("bigo")          
			Freportcontents     =rsget("reportcontents")
			Freportstate        =rsget("reportstate")    
			Fadminid            =rsget("adminid")        
		  Farap_nm        		=rsget("arap_nm")   
		  Facc_cd          		=rsget("acc_cd")   
		  Facc_use_cd					=rsget("acc_use_cd") 
		  Facc_nm         		=rsget("acc_nm")   
		   
		  Fpayrequestdate     =rsget("payrequestdate")   
		  Fpayrequestprice    =rsget("payrequestprice")   
		  FinBank            	=rsget("inBank")   
		  FaccountNo          =rsget("accountNo")   
		  FaccountHolder      =rsget("accountHolder")   
		  Fpaydate           	=rsget("paydate")   
		  FoutBank            =rsget("outBank") 
		  Fpayrealdate        =rsget("payrealdate") 
		   
		  Fyyyymm            	=rsget("yyyymm") 
		  FisTakeDoc          =rsget("isTakeDoc") 
		  Fpayrequeststate    =rsget("payrequeststate") 
		  Fregdate            =rsget("regdate")  
		  FpayComment					=rsget("comment")
		  Fusername						=rsget("username")
		  Fpartname						=rsget("part_name")
		  FpayRequestTitle		=rsget("payRequestTitle")
		 	Fcust_cd						=rsget("cust_cd")
		 	Fcust_nm						=rsget("cust_nm")
		 	Fbiz_no							=rsget("biz_no")
		 	Fpaytype 						=rsget("paytype")
		 	Fcurrencytype 			=rsget("currencytype")
		 	Fcurrencyprice			=rsget("currencyprice")
		 	FACC_GRP_CD					=rsget("ACC_GRP_CD")
		END IF
		rsget.close 
	END Function
	
		'//결제요청서 기본정보 내용보기
	public Function fnGetPayRequestReceiveData
	Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getReceiveData]( "&FpayrequestIdx&")"  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			Farap_cd			=rsget("arap_cd")
			FreportName         =rsget("reportName")    
			FreportPrice        =rsget("reportPrice")   
			Fscmlinkno          =rsget("scmlinkno")     
			Fbigo               =rsget("bigo")          
			Freportcontents     =rsget("reportcontents")
			Freportstate        =rsget("reportstate")    
			Fadminid            =rsget("adminid")        
		  Farap_nm        		=rsget("arap_nm")   
		  Facc_cd          		=rsget("acc_cd")   
		  Facc_use_cd				  =rsget("acc_use_cd")
		  Facc_nm         		=rsget("acc_nm")   
		  FedmsName           =rsget("edmsName")      
		  Fedmscode           =rsget("edmscode") 
		  FlastApprovalid     =rsget("lastApprovalid") 
		  
		  Fpayrequestdate     =rsget("payrequestdate")   
		  Fpayrequestprice    =rsget("payrequestprice")   
		  FinBank            	=rsget("inBank")   
		  FaccountNo          =rsget("accountNo")   
		  FaccountHolder      =rsget("accountHolder")   
		  Fpaydate           	=rsget("paydate")   
		  FoutBank            =rsget("outBank") 
		  Fpayrealdate        =rsget("payrealdate") 
		   
		  Fyyyymm            	=rsget("yyyymm") 
		  FisTakeDoc          =rsget("isTakeDoc") 
		  Fpayrequeststate    =rsget("payrequeststate") 
		  Fregdate            =rsget("regdate")  
		  FpayComment					=rsget("comment")
		  Fusername						=rsget("username")
		  Fpartname						=rsget("part_name")
		  FpayRequestTitle		=rsget("payRequestTitle") 
		  Fcust_cd						=rsget("cust_cd")
		  Fcust_nm						=rsget("cust_nm")
		  Freportidx					=rsget("reportidx")
		  FBiz_no							=rsget("biz_no")
		  FpayType						=rsget("payType")
		  FcurrencyType				=rsget("currencyType")
		  FcurrencyPrice			=rsget("currencyPrice")
		  FerpLinkType				=rsget("erpLinkType")
		  FACC_GRP_CD					= rsget("ACC_GRP_CD")
		END IF
		rsget.close 
	END Function
	
	'//기 결제요청서 리스트
	Function fnGetProcPayRequestList
		Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayRequest_getProcList("&FreportIdx&", "&FpayrequestIdx&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetProcPayRequestList = rsget.getRows()
		END IF
		rsget.close
	End Function
 
	
	'//품의서별 결제요청서 등록가능여부 확인
	Function fnCheckPayRequest
	Dim objCmd,returnValue
	 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText  		
				.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_chkReg]( "&FreportIdx&")}"							 
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With	
			    returnValue = objCmd(0).Value	
		Set objCmd = nothing 
	fnCheckPayRequest = returnValue 
	End Function
	
	'//결제요청서 받은리스트 가져오기
	public Function fnGetPayRequestReceiveList
	Dim strSql	 
		 IF Fpayrequeststate = "" THEN Fpayrequeststate = 1
		 IF Fauthstate = "" THEN Fauthstate = 0
		 IF FisLast = "" THEN FisLast = 1
		 	
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getReceiveListCnt]("&Fpayrequeststate&","&Fauthstate&","&FisLast&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayRequest_getReceiveList("&Fpayrequeststate&","&Fauthstate&","&FisLast&","&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayRequestReceiveList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'//품의서 결재라인 결제요청서 리스트
	public Function fnGetPayRequestAuthLine
	Dim strSql		
	strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getAuthListCnt]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayRequest_getAuthList('"&FadminID&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayRequestAuthLine = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'//결재 내용 수신확인 
	public Function fnCheckPayRequestView
	IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		Dim objCmd,returnValue 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_chkView]( "&FreportIdx&","&FpayrequestIdx&",'"&FadminId&"',"&Fauthposition&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
		Set objCmd = nothing 
		fnCheckPayRequestView = returnValue
	End Function
	
	public FPayRequeststate000 
	public FPayRequeststate001 
	public FPayRequeststate110 
	public FPayRequeststate111 
	public FPayRequeststate710 
	public FPayRequeststate711 
	public FPayRequeststate970 
	public FPayRequeststate971 
	public FPayRequeststate550 
	public FPayRequeststate551 

	'//재무회계 결제요청서 왼쪽 메뉴 카운트
	public Function fnGetLeftMenu
	Dim strSql		
		strSql ="[db_partner].[dbo].sp_Ten_eappPayRequest_receiveCount('"&FadminID&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FPayRequeststate000 = rsget("state000")
			FPayRequeststate001 = rsget("state001")
			FPayRequeststate110 = rsget("state110")
			FPayRequeststate111 = rsget("state111")
			FPayRequeststate710 = rsget("state710")
			FPayRequeststate711 = rsget("state711")
			FPayRequeststate970 = rsget("state970")
			FPayRequeststate971 = rsget("state971")
			FPayRequeststate550 = rsget("state550")
			FPayRequeststate551 = rsget("state551") 	
		END IF
		rsget.close
	 
	End Function
	
		'//증빙서류 데이터 가져오기
		public Function fnGetEappPayDoc
		Dim strSql		
		strSql ="db_partner.dbo.sp_Ten_eAppPayDoc_getData("&FPayRequestIdx&")"  
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FpayDocIdx  	= rsget("payDocIdx") 	
				Fpaydockind 	= rsget("paydockind")   
				Fvatkind  		= rsget("vatkind")  		
				Fissuedate  	= rsget("issuedate")    
				Fitemname  		= rsget("itemname")  	  
				Ftotprice 		= rsget("totprice") 		
				Fsupplyprice 	= rsget("supplyprice")  
				Fvatprice  		= rsget("vatprice")  	  
				Fetaxkey  		= rsget("etaxkey")  		
				FDocbigo  		= rsget("bigo")  			  
				Fattachfile 	= rsget("attachfile")
				FerpDocLinkType	= rsget("erpDocLinkType")
				FerpDocLinkKey	= rsget("erpDocLinkKey")
			END IF
				rsget.close
		End Function
End Class

'=====Function ==================================================================

Function fnGetPayRequestState(ByVal payrequeststate)
 DIM strMsg
	IF payrequeststate = 0 THEN	
		strMsg = "<font color='#777777'>작성중</font>"
	ELSEIF 	payrequeststate = 1 THEN
		strMsg = "결제진행중"
	ELSEIF 	payrequeststate = 7 THEN
		strMsg = "<font color='#3333FF'>결제승인</font>"
	ELSEIF 	payrequeststate = 5 THEN
		strMsg = "<font color='#FF33FF'>결제반려</font>"	
	ELSEIF 	payrequeststate = 8 THEN
		strMsg = "<font color='#11AA11'>ERP전송</font>"	
	ELSEIF 	payrequeststate = 9 THEN
		strMsg = "<font color='#FF3333'>결제완료</font>"			
	END IF
	fnGetPayRequestState = strMsg
End Function

Sub sbOptPayRequestState(ByVal payrequeststate)
%>
	<option value="0" <%IF payrequeststate="0" THEN%>selected<%END IF%>>작성중</option>
	<option value="1" <%IF payrequeststate="1" THEN%>selected<%END IF%>>결제진행중</option>
	<option value="7" <%IF payrequeststate="7" THEN%>selected<%END IF%>>결제승인</option>
	<option value="8" <%IF payrequeststate="8" THEN%>selected<%END IF%>>ERP전송</option>
	<option value="5" <%IF payrequeststate="5" THEN%>selected<%END IF%>>결제반려</option>
	<option value="9" <%IF payrequeststate="9" THEN%>selected<%END IF%>>결제완료</option>
	<option value="255" <%IF payrequeststate="255" THEN%>selected<%END IF%>>미완료전체</option>
<%	
End Sub
  
 Function fnGetPayAuthState(ByVal AuthState,ByVal AuthType)
 Dim strState
 Dim strWord
 	IF AuthType = 1 THEN
 		strWord ="승인"	
 	ELSE
 		strWord ="처리"	
 	END IF
  	IF AuthState =1 or AuthState  = 7 THEN  
		strState=strWord&"완료"
  	ELSEIF AuthState =3 THEN  
		strState=strWord&"보류"							
	ELSEIF AuthState =5 THEN
		strState=strWord&"반려"							
	ELSE	
		strState=strWord&"대기"						
	END IF
	fnGetPayAuthState = strState
 End Function
 
 
 '//결제방법 option
 Sub sboptPayType(ByVal ipaytype)
 %>
 <option value="0">--선택--</option>
 <option value="2" <%IF ipaytype="2" THEN%>selected<%END IF%>>계좌이체</option>
 <option value="1" <%IF ipaytype="1" THEN%>selected<%END IF%>>외화결제</option>
 <option value="3" <%IF ipaytype="3" THEN%>selected<%END IF%>>자동이체</option>
 <option value="4" <%IF ipaytype="4" THEN%>selected<%END IF%>>고지서납부</option>
 <option value="5" <%IF ipaytype="5" THEN%>selected<%END IF%>>Check대체</option>
 <option value="9" <%IF ipaytype="9" THEN%>selected<%END IF%>>기타결제</option>
 <%
End Sub

'//결제방법 값 
Function fnGetPayType(ByVal ipaytype)
Dim strPayType
	IF ipaytype="1" THEN 
		strPayType = "외화결제"
	ELSEIF ipaytype="2" THEN 
		strPayType = "계좌이체"
	ELSEIF ipaytype="3" THEN 
		strPayType = "자동이체"
	ELSEIF ipaytype="4" THEN 
		strPayType = "고지서납부"
	ELSEIF ipaytype="5" THEN 
		strPayType = "Check대체"
	ELSEIF ipaytype="9" THEN 
		strPayType = "기타결제"
	END IF   
	fnGetPayType = strPayType
End Function		
%>  