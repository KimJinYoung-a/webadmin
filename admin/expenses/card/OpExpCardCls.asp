<%
Class OpExp
	public FYYYYMM
	public FPartTypeIdx
	public FOpExpidx
	public FOpExpPartIdx
	public FOpExpDailyIdx
	public FYYYYMMDD 	 
	public Farap_cd	
	public Farap_nm
	public FinExp 			
	public FOutExp 		
	public FOpExpObj 		
	public FDetailCOnts 	
	public Fbizsection_cd
	public Fbizsection_nm
	public FsupExp 		
	public FvatExp 		
	public FauthNo 	 
  public FLastMonthExp
	public FTotExp
	public FOpExpPartName 
	
	public FSPageNo
	public FEPageNo
	public FPageSize
	public FCurrPage
	public FTotCnt 
	
	public FadminID
	public FSYYYYMM
	public FEYYYYMM
	public FPart_sn
	public FMode
	public FRectUserid
	public FRectPartsn
	public FState
		
	public Faccountidx
	public Finouttype
  public Fauthdate
  public FSAuthDate
  public FEAuthDate
  public FsevExp
  public Fdeducttype
	public FisYYYYMM
		
	'��� ����Ʈ
	public Function fnGetOpExpMonthlyList
		IF FPartTypeIdx ="" THEN FPartTypeIdx = 0
		IF FOpExpPartIdx ="" THEN FOpExpPartIdx = 0 
		IF FRectPartsn = "" THEN FRectPartsn = 0
		Dim strSql
		strSql = "[db_partner].[dbo].sp_Ten_OpExpMonthlyCard_getList('"&FSYYYYMM&"','"&FEYYYYMM&"',"&FOpExpPartIdx&",'"&FRectUserid&"',"&FRectPartsn&",'"&FState&"')"     
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpMonthlyList = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'��� �󼼸���Ʈ
	public Function fnGetOpExpDailyList  	 
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0 
 		IF Farap_cd = "" THEN Farap_cd = 0
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getListCnt]('"&FYYYYMM&"', "&FOpExpPartIdx&" ,"&Farap_cd&",'"&Fbizsection_nm&"')"     
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getList('"&FYYYYMM&"',  "&FOpExpPartIdx&","&Farap_cd&",'"&Fbizsection_nm&"',"&FSPageNo&","&FEPageNo&")"   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailyList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	
		
	'��� ��û��  �󼼸���Ʈ
	public Function fnGetOpExpDailyNoSetList  	 
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0 
 		IF Farap_cd = "" THEN Farap_cd = 0
 			IF FRectPartsn = "" THEN FRectPartsn = 0
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getNoSetListCnt]('"&FSAuthDate&"','"&FEAuthDate&"', "&FOpExpPartIdx&" ,"&Farap_cd&",'"&Fbizsection_nm&"','"&FisYYYYMM&"','"&FRectUserid&"',"&FRectPartsn&")"    
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getNoSetList('"&FSAuthDate&"','"&FEAuthDate&"',  "&FOpExpPartIdx&","&Farap_cd&",'"&Fbizsection_nm&"','"&FisYYYYMM&"','"&FRectUserid&"',"&FRectPartsn&","&FSPageNo&","&FEPageNo&")"   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailyNoSetList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'��� �� �հ� ����Ʈ
	public Function fnGetOpExpDailySumList
		Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getSumList('"&FYYYYMM&"',"&FOpExpPartIdx&","&Farap_cd&")"    
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailySumList = rsget.getRows()
		END IF
		rsget.close 
	End Function

	'��� ��������
	public Function fnGetOpExpDailyData
	Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getData]("&FOpExpDailyIdx&")"   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN 
			FYYYYMM 		= rsget("YYYYMM")
			Fauthdate		= rsget("authdate")
			FOpExpPartIdx 	= rsget("OpExpPartIdx")
			Farap_cd 		= rsget("arap_cd") 
			FOutExp 		= rsget("OutExp")
			FOpExpObj 		= rsget("OpExpObj")
			FDetailCOnts 	= rsget("DetailCOnts")
			Fbizsection_Cd= rsget("bizsection_cd")
			FsupExp 		= rsget("supExp")
			FvatExp 		= rsget("vatExp")
			FsevExp 		= rsget("sevExp")
			FauthNo 		= rsget("authNo")  
			Fdeducttype	= rsget("deducttype")
			Finouttype	= rsget("inouttype")
			Fbizsection_nm = rsget("bizsection_nm")
		END IF
		rsget.close
	End Function

	public Function fnGetOpExpMonthlyData
		Dim strSql	    
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
		IF FopExpIdx = "" THEN FopExpIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpMonthlyCard_getData]('"&Fyyyymm&"',"&FOpExpPartIdx&","&FopExpIdx&")"    
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN 
			FOpExpidx 		= rsget("OpExpCardidx") 
			Fyyyymm				= rsget("yyyymm") 
			FOpExpPartIdx	= rsget("OpExpPartIdx")  
			FOutExp				= rsget("OutExp")   
			FOpExpPartName=rsget("OpExpPartName") 
			Fstate       	= rsget("state")
		END IF
		rsget.close
	End Function
	
	
	'//����� ���� üũ
	public Function fnGetOpExpPartAuth  
	Dim objCmd
	 Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPart_getAuth]( "&FOpExpPartIdx&",'"&FadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    fnGetOpExpPartAuth = objCmd(0).Value	  
	Set objCmd = nothing	 
	End Function
End Class

'//���°�
Function fnGetStateDesc(ByVal iState)
	Dim strState
	IF iState = "1" THEN
		strState = "�ۼ��Ϸ�" 
	ELSEIF iState = "5" THEN
		strState = "����������"	
	ELSEIF iState = "7" THEN
		strState = "<font color='#3333FF'>����Ϸ�</font>"
	ELSEIF iState = "9" THEN
		strState = "<font color='#11AA11'>Ȯ�οϷ�</font>"	
	ELSEIF iState = "10" THEN
		strState = "<font color='#FF33FF'>���ۿϷ�</font>"	
	ELSE
		strState = "<font color='red'>�ۼ���</font>"	
	END IF		
	fnGetStateDesc = strState
End Function

Sub SbOptState(ByVal iState)
	%>
	<option value="">--����--</option>
	<option value="0" <%IF iState ="0" THEN%>selected<%END IF%>>�ۼ���</option>
	<option value="1" <%IF iState ="1" THEN%>selected<%END IF%>>�ۼ��Ϸ�</option>
	<option value="5" <%IF iState ="5" THEN%>selected<%END IF%>>����������</option>
	<option value="7" <%IF iState ="7" THEN%>selected<%END IF%>>����Ϸ�</option>
	<option value="9" <%IF iState ="9" THEN%>selected<%END IF%>>Ȯ�οϷ�</option>
	<option value="10" <%IF iState ="10" THEN%>selected<%END IF%>>���ۿϷ�</option>
	<%
End Sub

'//���Ѱ��� - ������ ����
Function fnChkAdminAuth( ByVal  authLevel, ByVal Partsn)
	Dim strAuth
	strAuth = False
	IF (authLevel<=2  or partsn= 8) THEN 
			strAuth = True
	END IF 
	fnChkAdminAuth = strAuth
End Function
 
%>
 