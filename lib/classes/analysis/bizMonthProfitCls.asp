<%
Class CBizProfit
public FYYYYMM
public Faccusecd
public Fbiztype
public FcolType
public FrowType
public FBizsection_Cd
public FAccGrpCd
public FSAccGrpCd
public FEAccGrpCd
'손익보고서
	public Function fnGetBizMonthProfitList
	 Dim strSql
	 IF Fbiztype = "" THEN Fbiztype = 0 '내부거래여부
	 IF FcolType = "" THEN FcolType = 1	'가로 보기옵션:1-사업부별, 2-팀별
	 IF FrowType = "" THEN FrowType = 1 '세로 보기옵션:1-계정그룹별, 2-계정구분별, 3- 계정과목별
	 	FSAccGrpCd = FAccGrpCd - 100 
	 	IF FAccGrpCd = "0"  THEN FSAccGrpCd = ""
	 strSql ="[db_partner].[dbo].sp_Ten_BizMonthProfit_getList('"&FYYYYMM&"','"&Faccusecd&"', "&Fbiztype&","&FcolType&","&FrowType&" ,'"&FSAccGrpCd&"' ,'"&FAccGrpCd&"')"	   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProfitList = rsget.getRows()
		END IF
		rsget.close
	End Function

'손익보고서 업무비율 구분	
	public Function fnGetBizMonthProfitBizList
	 Dim strSql
	 IF Fbiztype = "" THEN Fbiztype = 0 '내부거래여부 
	 strSql ="[db_partner].[dbo].sp_Ten_BizMonthProfitBiz_getList('"&FYYYYMM&"','"&Faccusecd&"', "&Fbiztype&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProfitBizList = rsget.getRows()
		END IF
		rsget.close
	End Function	 
	
'손익보고서 업무비율 구분상세
public Function fnGetBizMonthProfitBizDetail
 Dim strSql 
	 strSql ="[db_partner].[dbo].sp_Ten_BizMonthProfit_Bizsection_GetDetail('"&FYYYYMM&"','"&FBizsection_Cd&"','"&Faccusecd&"')"	   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProfitBizDetail = rsget.getRows()
		END IF
		rsget.close

End Function	
End Class
%>
