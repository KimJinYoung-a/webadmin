<%
Class CPayReqList
public FreportIdx
public FadminId
public FpayrequestIdx
public FaccountIdx			
public FreportName      
public FreportPrice     
public Fscmlinkno       
public Fbigo            
public Freportcontents  
public Freportstate     
public Freferid         
 
public Fregdate         
public FaccountName     
public Fcomm_name       
public Fcomm_desc       
public FerpCode         
public FedmsName        
public Fedmscode        
public FlastApprovalid  
public Fonline          
public Foffline         
public Fithinkso        
public Fbnw             
public Ffingers         
                 
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

public FpayRequestType  
 
public Farap_Cd
public FSearchType			 
public FSDate
public FEDate 
public FRegID		
public FBizSection_CD
public FnotIncEtc
public FcustNm
public FDocSendErp 
public FpayType
public Fpaydockind
public Fpayrequesttitle

	'//결제요청서 리스트 가져오기
	public Function fnGetPayReqAllList 
	Dim strSql	 
		IF FpayrequestIdx = "" then FpayrequestIdx = 0
		IF FpayRequestType = "" THEN FpayRequestType = 0
		IF FSearchType = "" THEN FSearchType = 1
		IF Farap_Cd = "" THEN Farap_Cd = 0 
		IF FpayRequestState = "" THEN FpayRequestState = -1
        IF Fpayrequestprice ="" THEN Fpayrequestprice =0
        ''if (FnotIncEtc="") THEN FnotIncEtc=""
        
		strSql ="[db_partner].[dbo].[sp_Ten_eappPayRequest_getAllListCnt]("&FpayrequestIdx&","&FpayRequestType&","&FSearchType&",'"&FSDate&"','"&FEDate&"' ,"&Farap_Cd&","&FpayRequestState&",'"&FisTakeDoc&"','"&Fusername&"',"&FOutBank&",'"&FBizSection_CD&"',"&Fpayrequestprice&",'"&FnotIncEtc&"','"&FcustNm&"','"&FDocSendErp&"',"&FpayType&",'"&Fpaydockind&"','"&FadminId&"','"&Fpayrequesttitle&"')"

		'rw strSql & "<br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eappPayRequest_getAllList("&FpayrequestIdx&","&FpayRequestType&","&FSearchType&",'"&FSDate&"','"&FEDate&"',"&Farap_Cd&","&FpayRequestState&",'"&FisTakeDoc&"','"&Fusername&"',"&FOutBank&",'"&FBizSection_CD&"',"&Fpayrequestprice&",'"&FnotIncEtc&"','"&FcustNm&"','"&FDocSendErp&"',"&FpayType&","&FSPageNo&","&FEPageNo&",'"&Fpaydockind&"','"&FadminId&"','"&Fpayrequesttitle&"')"

		'rw strSql & "<br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayReqAllList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
 End Class
 

Function PayDocKindName(code)
	SELECT CASE code
		Case "1" : PayDocKindName = "전자"
		Case "2" : PayDocKindName = "수기"
		Case "5" : PayDocKindName = "기타영수증"
		Case "8" : PayDocKindName = "차후 수취"
		Case "9" : PayDocKindName = "서류없음"
	CASE ELSE
		PayDocKindName = ""
	END SELECT
End Function
%>