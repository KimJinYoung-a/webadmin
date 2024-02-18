<%
 Class CAccount
public FaccountIdx  
public FaccountKind
public FaccountName
public FedmsIdx
public FIsUsing  
public Fregdate 

public Fcateidx1
public Fcateidx2

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage	
public FTotCnt
  

	'계정과목 리스트 가져오기
	public Function fnGetAccountList
		Dim strSql		
		IF FaccountKind = "" THEN FaccountKind = 0	
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAccount_getListCnt]("&FaccountKind&",'"&FaccountName&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppAccount_getList("&FaccountKind&",'"&FaccountName&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'계정과목 전자결재 리스트 가져오기
	public Function fnGetEappAccountList
		Dim strSql		
		IF FaccountKind = "" THEN FaccountKind = 0	
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAccount_getEappListCnt]("&FaccountKind&",'"&FaccountName&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppAccount_getEappList("&FaccountKind&",'"&FaccountName&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEappAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'공통코드내용 가져오기
	public Function fnGetAccountData
		Dim strSql		 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAccount_getData]( "&FaccountIdx&")"		
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			FaccountIdx 	= rsget("accountIdx")
			FaccountKind    = rsget("accountKind") 
			FedmsIdx        = rsget("edmsIdx")     
			FaccountName    = rsget("accountName") 
			Fregdate        = rsget("regdate") 
			Fcateidx1		= rsget("cateidx1")
			IF isNull(Fcateidx1) THEN   Fcateidx1 = 0  
			Fcateidx2		= rsget("cateidx2")    
			IF isNull(Fcateidx2) THEN   Fcateidx2 = 0
		END IF
		rsget.close
	End Function
 
    '특정문서에 해당하는 계정과목내용명 가져오기
    public Function fnGetEdmsAccountList
    	Dim strSql		
    	IF FedmsIdx = "" THEN FedmsIdx=0
		IF FaccountKind = "" THEN FaccountKind = 0	
		strSql ="[db_partner].[dbo].[sp_Ten_eAppAccount_getEdmsListCnt]("&FedmsIdx&","&FaccountKind&",'"&FaccountName&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppAccount_getEdmsList("&FedmsIdx&","&FaccountKind&",'"&FaccountName&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEdmsAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
 End Class
 
  
%>