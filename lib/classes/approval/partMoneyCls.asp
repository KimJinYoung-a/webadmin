<%	
 Class CpartMoneyCls 
 	public FeappDepth
 	public Fstep1partidx
 	public Fstep2partidx
 	public FeappPartName
 	public FisUsing
 	public FeappPartIdx
 	public FpartSort
 	
 	public FPageSize
 	public FCurrPage
 	public FSPageNo
 	public FEPageNo
    public FTotCnt 	
    
	'//list
	Function fnGetPartList 
		IF Fstep1partidx = "" THEN Fstep1partidx = 0
		IF Fstep2partidx = "" THEN Fstep2partidx = 0
		IF FeappDepth = "" THEN FeappDepth = 0
		Dim strSql	  
		strSql ="[db_partner].[dbo].sp_Ten_eappPart_getList("&Fstep1partidx&","&Fstep2partidx&","&FeappDepth&")"    
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPartList = rsget.getRows()
		END IF
		rsget.close 
	End Function
	
	'//get data
	Function fnGetPartData
		Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_eappPart_getData("&FeappPartIdx&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 FeappDepth		= rsget("eappDepth")
			 Fstep1partidx  = rsget("step1partidx")  
			 Fstep2partidx	= rsget("step2partidx")  
			 FeappPartName  = rsget("eappPartName")
			 FisUsing       = rsget("isUsing")    
		END IF
		rsget.close
	End Function 
	
	'//get 3depth List
	Function fnGet3DepthPartList
		Dim strSql	  
		strSql ="[db_partner].[dbo].sp_Ten_eappPart_get3DepthAllList "   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGet3DepthPartList = rsget.getRows()
		END IF
		rsget.close 
	End Function
	
	 Sub sb3DepthOptPart(ByVal iValue)
	 	if iValue = "" then iValue = 0
	 	Dim arrList, intLoop
	 	arrList = fnGet3DepthPartList
	 	IF isArray(arrList) THEN
	 		For intLoop = 0 To UBound(arrList,2)
 %>
 	<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(iValue) = Cstr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
 <%			Next
 		END IF
	End Sub
 End Class 
%>