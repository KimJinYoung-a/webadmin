<%
'############################
' Description : ������ Ŭ����
' History : 2011.04.21 ������  ����
'############################

Class COpExpAccount
public Farap_nm
public Farap_cd
public FOpExpPartIdx

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage
public FTotCnt

public FARAP_GB
public FCASH_FLOW
public frectarap_nm

	'��� �������� ����Ʈ ��������
	public Function fnGetOpExpAccountList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccount_getListCnt]('"&Farap_nm&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_OpExpAccount_getList('"&Farap_nm&"',"&FSPageNo&","&FEPageNo&")"	 
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
		strSql ="[db_partner].[dbo].[sp_Ten_TMS_BA_ARAP_CD_getOpExpListcNT]('"&Farap_nm&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_ARAP_CD_getOpExpList('"&Farap_nm&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'��� �����׸� ���� ����Ʈ
	public Function fnGetArapRegList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccount_getRegList]("&FOpExpPartIdx&",'"& frectarap_nm &"')"	

		'response.write strSql & "<Br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetArapRegList = rsget.getRows()
		END IF
		rsget.close
	End Function
	
		public Function fnGetArapAllList
		Dim strSql	 
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccount_getAllList]"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetArapAllList = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'���������� �ش��ϴ� �������� ����Ʈ
	public Function fnGetAccountData
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpAccountData_getList]("&Farap_cd&")"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetAccountData = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'��� ���ޱݾ� �����׸� ����Ʈ
	Function fnGetArapOutList
	Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpexpARAP_getOutList]('"&FARAP_GB&"','"&FCASH_FLOW&"','"&FARAP_NM&"')"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetArapOutList = rsget.getRows()
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