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
'���ͺ���
	public Function fnGetBizMonthProfitList
	 Dim strSql
	 IF Fbiztype = "" THEN Fbiztype = 0 '���ΰŷ�����
	 IF FcolType = "" THEN FcolType = 1	'���� ����ɼ�:1-����κ�, 2-����
	 IF FrowType = "" THEN FrowType = 1 '���� ����ɼ�:1-�����׷캰, 2-�������к�, 3- ��������
	 	FSAccGrpCd = FAccGrpCd - 100 
	 	IF FAccGrpCd = "0"  THEN FSAccGrpCd = ""
	 strSql ="[db_partner].[dbo].sp_Ten_BizMonthProfit_getList('"&FYYYYMM&"','"&Faccusecd&"', "&Fbiztype&","&FcolType&","&FrowType&" ,'"&FSAccGrpCd&"' ,'"&FAccGrpCd&"')"	   
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProfitList = rsget.getRows()
		END IF
		rsget.close
	End Function

'���ͺ��� �������� ����	
	public Function fnGetBizMonthProfitBizList
	 Dim strSql
	 IF Fbiztype = "" THEN Fbiztype = 0 '���ΰŷ����� 
	 strSql ="[db_partner].[dbo].sp_Ten_BizMonthProfitBiz_getList('"&FYYYYMM&"','"&Faccusecd&"', "&Fbiztype&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProfitBizList = rsget.getRows()
		END IF
		rsget.close
	End Function	 
	
'���ͺ��� �������� ���л�
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
