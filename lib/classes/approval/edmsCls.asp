<%
 Class Cedms
 public Fcatedepth
 public Fpcateidx
 public Fcategoryidx
 public Fcatename
 public Fcatecode
 public Fregdate

public FcateIdx1
public FcateIdx2

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage
 public FTotCnt

 public  Fedmsidx
 public FserialNum
 public Fedmsname
 public Fedmscode
 public FviewNo
 public FedmsFile
 public FisApproval
 public FisScmApproval
 public FlastApprovalid
 public FscmLink
 public FscmsubmitLink
 public Fadminid
 public FisUsing

 public FPayEApp
 public Fedmsform
 public FCfoAgree
 public FisAgreeNeed
 public FisAgreeNeedTarget
 public FisAgreeNeedTargetName

 	'카테고리 리스트 가져오기
 	public Function fnGetedmsCategoryList
 		Dim strSql
 		IF Fcatedepth = "" THEN Fcatedepth = 1
		IF Fpcateidx = "" THEN Fpcateidx = 0
		FTotCnt = 0
 		strSql ="db_partner.dbo.sp_Ten_edms_category_getList("&Fcatedepth&","&Fpcateidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 fnGetedmsCategoryList = rsget.getRows()
			 FTotCnt = ubound(fnGetedmsCategoryList,2)+1
		END IF
		rsget.close
	End Function

	'카테고리 정보 가져오기
	public Function fnGetedmsCategoryData
		Dim strSql
		strSql = "db_partner.dbo.sp_Ten_edms_category_getData("&Fcategoryidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			  Fcatedepth= rsget("catedepth")
			  Fcatename	= rsget("catename")
			  Fcatecode = rsget("catecode")
			  Fpcateidx = rsget("pcateidx")
			  Fregdate	= rsget("regdate")
		END IF
		rsget.close
	End Function

	'카테고리 depth별 select-box 옵션리스트로 가져오기
	public Sub sbGetOptedmsCategory(ByVal catedepth, ByVal pcateidx, ByVal cateidx)
		Dim arrList ,intLoop
		Fcatedepth = catedepth
		Fpcateidx = pcateidx

		arrList = fnGetedmsCategoryList
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(cateidx) =  Cstr(arrList(0,intLoop))  THEN%>selected<%END IF%>><%=arrList(3,intLoop)%>-<%=arrList(2,intLoop)%></option>
	<%		Next
		END IF
	End Sub

	'중카테고리 코드 자동생성
	public Function fnGetCatecode
		Dim strSql
		strSql = "db_partner.dbo.sp_Ten_edms_category_getCode("&Fpcateidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			  fnGetCatecode = rsget("catecode")
		END IF
		rsget.close
	End Function

	'문서리스트 가져오기
	public Function fnGetEdmsList
		Dim strSql

		strSql ="[db_partner].[dbo].[sp_Ten_edms_getListCnt]("&FCateIdx1&" ,"&FCateIdx2&",'"&Fedmsname&"','"&FisUsing&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_edms_getList("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"','"&FisUsing&"',"&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEdmsList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'전자결재 문서리스트 가져오기
	public Function fnGetEappEdmsList
		Dim strSql

		strSql ="[db_partner].[dbo].[sp_Ten_edms_getEappListCnt]("&FCateIdx1&" ,"&FCateIdx2&",'"&Fedmsname&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_edms_getEappList("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"',"&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEappEdmsList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'계정과목 선택용  문서리스트(결제요청서 연동 유)  가져오기
	public Function fnGetPayEdmsList
		Dim strSql

		strSql ="[db_partner].[dbo].[sp_Ten_edms_getPayListCnt]("&FCateIdx1&" ,"&FCateIdx2&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_edms_getPayList("&FCateIdx1&","&FCateIdx2&","&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayEdmsList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'문서내용 가져오기
	public Function fnGetEdmsData
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_edms_getData]( "&Fedmsidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			Fedmsidx       	= rsget("edmsidx")
			Fcateidx1       = rsget("cateidx1")
			Fcateidx2       = rsget("cateidx2")
			FserialNum    	= rsget("serialNum")
			Fedmsname  		= rsget("edmsname")
			Fedmscode   	= rsget("edmscode")
			FviewNo         = rsget("viewNo")
			FedmsFile      	= rsget("edmsFile")
			FisApproval    	= rsget("isApproval")
			FisScmApproval  = rsget("isScmApproval")
			FlastApprovalid = rsget("lastApprovalid")
			FscmLink       	= rsget("scmLink")
			FscmsubmitLink 	= rsget("scmsubmitLink")
			Fregdate        = rsget("regdate")
			Fadminid        = rsget("adminid")
			FisUsing		= rsget("isUsing")
			FPayEApp		= rsget("isPayEApp")
			Fedmsform		= replace(nl2blank(rsget("edmsform")),"'","\'")
			FCfoAgree       = rsget("CfoAgree")
			FisAgreeNeed	= rsget("isAgreeNeed")
			FisAgreeNeedTarget = rsget("isAgreeNeedTarget")
			FisAgreeNeedTargetName = rsget("username")
		END IF
		rsget.close
	End Function


	'중카테고리 일련번호 자동생성
	public Function fnGetSerialNum
		Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_edms_getSerialNum]("&Fcateidx1&","&Fcateidx2&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			  fnGetSerialNum = rsget("serialnum")
		END IF
		rsget.close
	End Function

	'문서리스트 select-box option형
	public Sub sbOptEdmsList
		Dim arrList ,intLoop
		FCurrPage = 1
		FPageSize = 100
		arrList = fnGetEdmsList
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(Fedmsidx) =  Cstr(arrList(0,intLoop))  THEN%>selected<%END IF%>><%=arrList(7,intLoop)%>-<%=arrList(6,intLoop)%></option>
	<%		Next
		END IF
	End Sub

	'문서리스트 select-box option형
	public Sub sbOptPayEdmsList
		Dim arrList ,intLoop
		FCurrPage = 1
		FPageSize = 100
		arrList = fnGetPayEdmsList
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
	%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(Fedmsidx) =  Cstr(arrList(0,intLoop))  THEN%>selected<%END IF%>><%=arrList(7,intLoop)%>-<%=arrList(6,intLoop)%></option>
	<%		Next
		END IF
	End Sub
 End Class
%>