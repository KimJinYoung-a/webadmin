<%
Class COpExpAccount
public Fcomm_name
public Fcomm_cd

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage
public FTotCnt

	'��� �������� ����Ʈ ��������
	public Function fnGetOpExpAccountList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccount_getListCnt]('"&Fcomm_name&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_OpExpAccount_getList('"&Fcomm_name&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'�������� ��ü ����Ʈ
	public Function fnGetAccountList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_eApppCommCD_getOpExpListCnt]('"&Fcomm_name&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eApppCommCD_getOpExpList('"&Fcomm_name&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'��� �������� ��ü ����Ʈ
	public Function fnGetAccountAll
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccount_getAllList]"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountAll = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'���������� �ش��ϴ� �������� ����Ʈ
	public Function fnGetAccountData
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccountData_getList]("&Fcomm_cd&")"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountData = rsget.getRows()
		END IF
		rsget.close
	End Function
End Class

	'��� �������� option ��
	Sub sbOptAccount(arrList,iValue)
	if isnull(iValue) then iValue = ""
	Dim intLoop
		If isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			%>
			<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(arrList(0,intLoop)) = Cstr(iValue) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
			<%
			Next
		END IF
	End Sub
%>