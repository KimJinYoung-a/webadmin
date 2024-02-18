<%
 Class CAccCategory
 public FACCDepth
 public FACCPCateIdx
 public FACCCateIdx
 public FACCCD
 public FACCCateName
 public FACCOrder
 public FisUsing
 public FACCUseCD
 public FACCNM	  
 public FisNoSet
 
 public FSPageNo
 public FEPageNo
 public FPageSize
 public FCurrPage	
 public FTotCnt

 public FSale10x10
 public FSalePartner
 public FDivide
 public FDividedesc

 	'카테고리 리스트 가져오기
 	public Function fnGetAccCategoryList
 		Dim strSql
 		IF FACCDepth = "" THEN FACCDepth = 1
		IF FACCPCateIdx = "" THEN FACCPCateIdx = 0  
		 
 		strSql ="db_partner.dbo.sp_Ten_ACC_CD_category_getList("&FACCDepth&","&FACCPCateIdx&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 fnGetAccCategoryList = rsget.getRows() 
		END IF
		rsget.close	
	End Function
	
	'카테고리 정보 가져오기
	public Function fnGetAccCategoryData
		Dim strSql
		strSql = "db_partner.dbo.sp_Ten_ACC_CD_Category_getData("&FACCCateIdx&")"	
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			  FACCCateName	= rsget("ACC_CateName")
			  FACCDepth 		= rsget("Acc_Depth")
			  FACCPCateIdx 	= rsget("Acc_PCateIdx")
			  FACCOrder 		= rsget("ACC_Order")
			  FisUsing 			= rsget("isUsing") 
		END IF
		rsget.close	
	End Function
	
	
	'카테고리 depth별 select-box 옵션리스트로 가져오기 
	public Sub sbGetOptACCCategory(ByVal catedepth, ByVal pcateidx, ByVal cateidx)
		Dim arrList ,intLoop
		FACCDepth = catedepth
		FACCPCateIdx = pcateidx  
		 
		arrList = fnGetAccCategoryList
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(cateidx) =  Cstr(arrList(0,intLoop))  THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
	<%		Next 
		END IF 
	End Sub	
	
	 
	'계정과목 카테고리 리스트 가져오기
	public Function fnGetACCCDList
		Dim strSql	
		IF FACCPCateIdx = "" THEN FACCPCateIdx = 0 
		IF FACCCateIdx = "" THEN FACCCateIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_ACC_CD_CategoryDetail_getListCnt]("&FACCPCateIdx&" ,"&FACCCateIdx&",'"&FACCUSECD&"','"&FACCNM&"','"&FisNoSet&"')"	 
	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_ACC_CD_CategoryDetail_getList("&FACCPCateIdx&" ,"&FACCCateIdx&",'"&FACCUSECD&"','"&FACCNM&"','"&FisNoSet&"',"&FSPageNo&","&FEPageNo&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetACCCDList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function 
	

	'계정과목 안분기준 내용 가져오기
	public Function fnGetACCDivData
	dim strSql
	strSql ="[db_partner].[dbo].sp_Ten_Acc_CD_CategoryDetail_getdivide('"&FACCCD&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN 
			FACCNM = rsget("acc_nm")
			FSale10x10 = rsget("issale10x10") 
			FSalePartner = rsget("issalepartner") 
			FDivide = rsget("divide")
			FDividedesc = rsget("dividedesc")
		END IF
		rsget.close
	 
	End Function
 End Class
%>