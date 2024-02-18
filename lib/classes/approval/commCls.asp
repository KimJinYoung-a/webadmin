<%
 Class CcommCode 
public Fcomm_cd  
public Fparentkey
public Fcomm_name
public Fcomm_desc
public FerpCode  
public Fdispnum  
public FactiveYN 
public FRectParentKey
 
public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage	
public FTotCnt
  

	'공통코드리스트 가져오기
	public Function fnGetCommCDList
		Dim strSql		
			
		strSql ="[db_partner].[dbo].[sp_Ten_eAppCommCD_getListCnt]("&Fparentkey&")"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eAppCommCD_getList("&Fparentkey&","&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCommCDList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
	
	'공통코드내용 가져오기
	public Function fnGetCommCDData
		Dim strSql		 
		strSql ="[db_partner].[dbo].[sp_Ten_eAppCommCD_getData]( "&Fcomm_cd&")"		
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			Fcomm_cd       	= rsget("comm_cd")
			Fparentkey      = rsget("parentkey")
			Fcomm_name      = rsget("comm_name")
			Fcomm_desc    	= rsget("comm_desc")
			FerpCode  		= rsget("erpCode")
			Fdispnum   		= rsget("dispnum")
			FactiveYN       = rsget("activeYN") 
		END IF
		rsget.close
	End Function
 
 	'특정그룹에 해당하는 공통코드 리스트 가져오기
 	public Function fnGetUseCommCDList
 	Dim strSql
 	strSql ="[db_partner].[dbo].[sp_Ten_eAppCommCD_getUseList]("&Fparentkey&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetUseCommCDList = rsget.getRows()
		END IF
		rsget.close
		 
	End Function
	
	
 	'공통코드 최상위행 가져오기
 	public Function fnGetCommCDGroup
 		Dim strSql
 		strSql ="[db_partner].[dbo].sp_Ten_eAppCommCD_getGroupList"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCommCDGroup = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'공통코드 최상위행(그룹) select-box option형
	public Sub sbOptCommCDGroup
		Dim arrList, intLoop
		arrList = fnGetCommCDGroup
		IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(FRectParentKey) = Cstr(arrList(0,intLoop)) THEN%> selected <%END IF%>><%=arrList(1,intLoop)%></option>
	<%	
		Next
		END IF
	End Sub
	
	'공통코드 리스트  select-box option형
	public Sub sbOptCommCD 
		Dim arrList, intLoop 
		arrList = fnGetUseCommCDList
		IF isnull(Fcomm_cd) THEN Fcomm_cd = ""
		IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(Fcomm_cd) = Cstr(arrList(0,intLoop)) THEN%> selected <%END IF%>><%=arrList(1,intLoop)%></option>
	<%	
		Next
		END IF
	End Sub
 End Class
 
  
%>