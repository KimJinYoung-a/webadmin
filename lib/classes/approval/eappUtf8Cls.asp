<%
Class CEApproval

public FedmsIdx    
public FaccountName
public Fedmsname   
public Fedmscode	
public Fcomm_name	
public FerpCode	
public FlastApprovalid
public Fjob_name
public Fdepartment_id
public Fdepartmentnamefull
public FscmLink
public FscmsubmitLink
public Fcomm_desc 
public Fpart_sn
public Fpart_name
public Fusername

public Fcid1
public Fcid2
public Fcid3
public Fcid4

public FreportName      
public FreportPrice     
public Fscmlinkno       
public Fbigo            
public Freportcontents   
public Freferid         
     
public Fregdate           
   

public FadminId
public FreportState
public FAuthState
public FreportIdx 

public FpayrequestIdx

public FPageSize
public FCurrPage
public FSPageNo
public FEPageNo
public FTotCnt
  
public Farap_cd
public Farap_nm
public Facc_cd
public Facc_nm
public Facc_use_cd
public FedmsForm 
public FisPayEapp
public Fpayrequestprice  
public FACC_GRP_CD
public FisAgreeNeed
public FisAgreeNeedTarget

public Fpaytype			
public Fcurrencytype	
public Fcurrencyprice

	'//결재 기본 폼 가져오기
	public Function fnGetEAppForm
		Dim strSql		
		IF  Farap_cd = "" THEN Farap_cd = 0
		IF  FedmsIdx = "" THEN FedmsIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppForm_getData]("&Farap_cd&", "&FedmsIdx&")"		
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			FedmsIdx        = rsget("edmsIdx")   
			Fedmsname       = rsget("edmsname") 
			Fedmscode				= rsget("edmscode") 
			FlastApprovalid = rsget("lastApprovalid")
			FscmLink   			= rsget("scmLink")
			Fjob_name				= rsget("job_name") 
			Farap_cd 				= rsget("arap_cd")
			Farap_nm    		= rsget("arap_nm")  
			Facc_cd    			= rsget("acc_cd")  
			Facc_nm					= rsget("acc_nm")   
			Facc_use_cd			= rsget("acc_use_cd") 
			FedmsForm				= replace(nl2blank(rsget("edmsform")),"'","\'")  
			FisPayEapp 	= rsget("isPayEapp")
			FACC_GRP_CD			= rsget("ACC_GRP_CD")
			FisAgreeNeed		= rsget("isAgreeNeed")
			FisAgreeNeedTarget	= rsget("isAgreeNeedTarget")
		END IF
		rsget.close
	End Function
	 
	'//보낸 결재함 리스트
	public Function fnGetEAppSendList
		Dim strSql	 
		IF FreportState = ""  THEN FreportState = -1
		strSql ="[db_partner].[dbo].[sp_Ten_eAppReport_getSendListCnt]('"&FadminId&"',"&FreportState&")"  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppReport_getSendList('"&FadminId&"',"&FreportState&","&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEAppSendList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'//받은결재함 리스트
	public Function fnGetEAppReceiveList
		Dim strSql	 
IF FAuthState = "" THEN FAuthState = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppReport_getReceiveListCnt]('"&FadminId&"',"&FAuthState&","&FreportState&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppReport_getReceiveList('"&FadminId&"',"&FAuthState&","&FreportState&","&FSPageNo&","&FEPageNo&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEAppReceiveList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'//참조 리스트
	public Function fnGetEAppReferList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppReport_getReferListCnt]('"&FadminId&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppReport_getReferList('"&FadminId&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEAppReferList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'//전자결재 기본정보 내용보기
	public Function fnGetEAppData
	Dim strSql	 
	IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppReport_getData]( "&FreportIdx&", "&FpayrequestIdx&")"		 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			Farap_cd						=rsget("arap_cd")
			FreportName         =rsget("reportName")    
			FreportPrice        =rsget("reportPrice")   
			Fscmlinkno          =rsget("scmlinkno")     
			Fbigo               =rsget("bigo")          
			Freportcontents     =replace(nl2blank(rsget("reportcontents")),"'","\'") 
			Freportstate        =rsget("reportstate")   
			Freferid            =rsget("referid")       
			Fadminid            =rsget("adminid")       
			Fregdate            =rsget("regdate")       
		  Farap_nm        		=rsget("arap_nm")   
		  Facc_cd          		=rsget("acc_cd")   
		  Facc_nm          		=rsget("acc_nm")       
		  FedmsName           =rsget("edmsName")      
		  Fedmscode           =rsget("edmscode") 
		  FlastApprovalid     =rsget("lastApprovalid")
		  IF FlastApprovalid = "0" THEN FlastApprovalid = rsget("eapp_lastApprovalid")
		  FscmLink						=rsget("scmLink")
		  FscmsubmitLink			=rsget("scmsubmitlink")
		  Fjob_name						=rsget("job_name")
		  Fdepartment_id			=rsget("department_id")
		  Fdepartmentnamefull	=rsget("departmentnamefull")
		  Fusername						=rsget("username")
		  FisPayEapp					=rsget("isPayEapp")
		  Fpayrequestprice		=rsget("payrequestprice")
		  Facc_use_cd					=rsget("acc_use_cd")
		  Fpaytype						=rsget("paytype")
		  Fcurrencytype				=rsget("currencytype")
		  Fcurrencyprice			=rsget("currencyprice")
		  FACC_GRP_CD					= rsget("ACC_GRP_CD")
		  Fcid1								= rsget("cid1")
		  Fcid2								= rsget("cid2")
		  Fcid3								= rsget("cid3")
		  Fcid4								= rsget("cid4")
		  FisAgreeNeed						= rsget("isAgreeNeed")
		  FisAgreeNeedTarget				= rsget("isAgreeNeedTarget")	
		END IF
		rsget.close 
	END Function

	'//결재라인
	public Function fnGetAuthLineList
		Dim strSql	 
		IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAuthLine_getData]( "&FreportIdx&", "&FpayrequestIdx&")"	
		'response.write strSql
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetAuthLineList =  rsget.getRows()
		END IF
		rsget.close 
	End Function	
	
	'//결재라인 - 반려리스트
	public Function fnGetAuthLineReturnList
		Dim strSql	 
		IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAuthLine_getReturnList]( "&FreportIdx&", "&FpayrequestIdx&")"		
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetAuthLineReturnList =  rsget.getRows()
		END IF
		rsget.close 
	End Function	
	
	
	'//첨부파일
	public Function fnGetAttachFileList
	Dim strSql	 
	IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAttachFile_getData]( "&FreportIdx&", "&FpayrequestIdx&")"  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetAttachFileList =  rsget.getRows()
		END IF
		rsget.close 
	End Function
	
	'//부서별 자금구분
	public Function fnGetPartMoneyList
	Dim strSql
	IF FpayrequestIdx = ""  THEN FpayrequestIdx = 0
	IF FpayrequestIdx < 0  THEN FpayrequestIdx = 0
			
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPartMoney_getList]( "&FreportIdx&", "&FpayrequestIdx&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetPartMoneyList =  rsget.getRows()
		END IF
		rsget.close 
	End Function
	
	'//코멘트
	public Function fnGetCommentList
	Dim strSql	 
	IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_eAppComment_getData]( "&FreportIdx&", "&FpayrequestIdx&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetCommentList =  rsget.getRows()
		END IF
		rsget.close 
	End Function
	
	'//결재 내용 수신확인 
	public Function fnCheckView
	IF FpayrequestIdx = "" THEN FpayrequestIdx = 0
		Dim objCmd,returnValue 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_chkView]( "&FreportIdx&","&FpayrequestIdx&",'"&FadminId&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
		Set objCmd = nothing 
		fnCheckView = returnValue
	End Function
	
public FReportstate0
public FReportstate1
public FReportstate3
public FReportstate5
public FReportstate7
public FReportstate100 
public FReportstate110 
public FReportstate710 
public FReportstate130 
public FReportstate150 
public FReportstate101 
public FReportstate111 
public FReportstate711 
public FReportstate131 
public FReportstate151 
public FPayRequeststate0
public FPayRequeststate9 
public FPayRequeststate1 
public FPayRequeststate5 
public FPayRequeststate7 
public Frefercount
public Fauthcount
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
public FDocCount

	'//왼쪽 메뉴 카운트
	public Function fnGetLeftMenu
		Dim strSql	
		'보낸결재함 > 결재문서 count  
		strSql ="[db_partner].[dbo].[sp_Ten_eappReport_sendCount]('"&FadminId&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
		 	FReportstate0 = rsget("state0")
		 	FReportstate1 = rsget("state1")
		 	FReportstate3 = rsget("state3")
		 	FReportstate5 = rsget("state5")
		 	FReportstate7 = rsget("state7")  
		END IF
		rsget.close 
		'받은결재함 > 결재문서 count
		strSql ="[db_partner].[dbo].[sp_Ten_eappReport_receiveCount]('"&FadminId&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
		 	FReportstate100 = rsget("state100")
		 	FReportstate110 = rsget("state110")
		 	FReportstate710 = rsget("state710")
		 	FReportstate130 = rsget("state130")
		 	FReportstate150 = rsget("state150")  
		 	FReportstate101 = rsget("state101")
		 	FReportstate111 = rsget("state111")
		 	FReportstate711 = rsget("state711")
		 	FReportstate131 = rsget("state131")
		 	FReportstate151 = rsget("state151")  
		END IF
		rsget.close 
		'보낸결재함 > 결제요청서 count
		strSql ="[db_partner].[dbo].[sp_Ten_eappPayRequest_sendCount]('"&FadminId&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			FPayRequeststate0 = rsget("state0") 
		 	FPayRequeststate1 = rsget("state1") 
		  	FPayRequeststate5 = rsget("state5")
		  	FPayRequeststate7 = rsget("state7")
		  	FPayRequeststate9 = rsget("state9")
		END IF
		rsget.close 
		
		'보낸결재함 > 결재문서> 참조 count 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppReport_getReferListCnt]('"&FadminId&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			Frefercount = rsget(0)
		END IF
		rsget.close
		
		'받은결재함 > 결제요청서>결재선 count
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getAuthListCnt]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FauthCount = rsget(0)
		END IF
		rsget.close
		
		'받은결재함 > 결제요청서 count
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
		
		'세금계산서 차후수취 count
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayDoc_getListCnt('"&FadminID&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FDocCount = rsget(0)
		END IF
		rsget.close	
	End Function
	
	public FTotSendCount
	public FTotReceiveCount
	public FTotReceiveViewCount
	public FTotpaySendCount
	
	'메인 카운트 
	public Function fnGetMainCount
		Dim strSql
		'//보낸결재함
		strSql ="[db_partner].[dbo].[sp_Ten_eappreport_TotSendCount]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotSendCount = rsget(0)
		END IF
		rsget.close
		'받은결재함
		strSql ="[db_partner].[dbo].[sp_Ten_eappreport_TotReceiveCount]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotReceiveCount = rsget("count0")
			FTotReceiveViewCount = rsget("count1")
		END IF
		rsget.close
		
		'//결제요청서
		strSql ="[db_partner].[dbo].[sp_Ten_eappPayRequest_TotSendCount]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotpaySendCount = rsget(0)
		END IF
		rsget.close 
		
		'받은결재함 > 결제요청서>결재선 count
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getAuthListCnt]('"&FadminID&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FauthCount = rsget(0)
		END IF
		rsget.close
	End Function
	
	'//결재승인목록에서 결제요청서 리스트가져오기
	public Function fnGetPayRequestReportList
	Dim strSql
		 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppPayRequest_getReportList]("&Freportidx&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayRequestReportList = rsget.getRows()
		END IF
		rsget.close
	
	End Function
End Class

'//메뉴명
Function fnGetMenu(ByVal menutype, ByVal reportstate, ByVal authstate)
	Dim strMenu
	IF menutype = "S1" THEN
		IF reportstate = "0" THEN
			strMenu = "작성중(임시저장)"
		ELSEIF reportstate ="1" THEN
			strMenu = "진행문서"
		ELSEIF reportstate ="7" THEN
			strMenu = "승인문서"
		ELSEIF reportstate ="3" THEN
			strMenu = "보류문서"
		ELSEIF reportstate ="5" THEN
			strMenu = "반려문서"
		ELSE
			strMenu = "신규작성"
		END IF 
	ELSEIF menutype = "S2" THEN
		IF reportstate = "0" THEN
			strMenu = "작성중(임시저장)"
		ELSEIF reportstate = "1" THEN
			strMenu = "결제요청"
		ELSEIF reportstate ="7" THEN
			strMenu = "결제승인"
		ELSEIF reportstate ="9" THEN
			strMenu = "결제완료" 
		ELSEIF reportstate ="5" THEN
			strMenu = "결제반려" 
		END IF 	
	ELSEIF menutype ="R1" THEN
		IF reportstate = "1" and authstate = "0" THEN
			strMenu = "결재대기"
		ELSEIF reportstate ="1" and authstate = "1" THEN
			strMenu = "결재완료(진행중)"
		ELSEIF reportstate ="7" and authstate = "1" THEN
			strMenu = "결재완료(최종승인)"
		ELSEIF (reportstate ="3" or reportstate ="1") and authstate = "3" THEN
			strMenu = "결재보류"
		ELSEIF (reportstate ="5"or reportstate ="1") and authstate = "5" THEN
			strMenu = "결재반려"
		ELSE
			strMenu = "참조"
		END IF 
	ELSEIF menutype ="R2" THEN	
			strMenu = "결재선"
	ELSEIF menutype ="FR" THEN	'재무회계팀 결제요청서 처리 메뉴
		IF reportstate = "1"  and authstate = 0 THEN
			strMenu = "결제요청전승인"
		ELSEIF reportstate = "1" and authstate = 1 THEN
			strMenu = "결제요청"
		ELSEIF reportstate = "7" THEN		
			strMenu = "결제확인(결제예정)"
		ELSEIF reportstate = "9" THEN
			strMenu = "결제완료"
		ELSEIF reportstate = "5" THEN
			strMenu = "결제반려"
		END IF		
	END IF	
	 
	fnGetMenu = "<font color=#4E9FC6>"&strMenu&"</font>"
End Function

 
 Function fnGetAuthState(ByVal AuthState)
 Dim strState
  	IF AuthState =1 THEN  
		strState="승인완료"
  	ELSEIF AuthState =3 THEN  
		strState="보류"							
	ELSEIF AuthState =5 THEN  
		strState="반려"							
	ELSE	
		strState="승인대기"						
	END IF
	fnGetAuthState = strState
 End Function
 
Function fnGetReportState(ByVal ReportState)
Dim strT
		IF reportstate ="1" THEN
			strT = "<font color=green>진행중</font>"
		ELSEIF reportstate ="7" or   reportstate ="8" or   reportstate ="9" THEN
			strT = "<font color=blue>승인</font>"
		ELSEIF reportstate ="3" THEN
			strT = "<font color=gray>보류</font>"
		ELSEIF reportstate ="5" THEN
			strT = "<font color=red>반려</font>" 
		ELSE
			strT = "작성중" 	
		END IF 
		
		fnGetReportState = strT
End Function

Sub sbOptReportState(ByVal ReportState)
%>
	<option value="0" <%IF ReportState="0" THEN%>selected<%END IF%>>작성중</option>
	<option value="1" <%IF ReportState="1" THEN%>selected<%END IF%>>진행중</option>
	<option value="7" <%IF ReportState="7" THEN%>selected<%END IF%>>승인</option>
	<option value="3" <%IF ReportState="3" THEN%>selected<%END IF%>>보류</option>
	<option value="5" <%IF ReportState="5" THEN%>selected<%END IF%>>반려</option> 
<%	
End Sub
%> 