<%
'###########################################################
' Description : �����α�_PG�纰 Ŭ����
' Hieditor : 2013.12.27 ������ ����
'###########################################################

Class CPGLog
	public Fdatetype
	public Fstartdate
	public Fenddate
	public Fpggubun
	public Fpguserid
	public FRectGroupBy

	public Function fnGetPGLogList
		Dim strSql
		strSql = "db_datamart.dbo.sp_Ten_order_payment_log_getPG("&Fdatetype&",'"&Fstartdate&"','"&Fenddate&"','"&Fpggubun&"','"&Fpguserid&"','"&FRectGroupBy&"')"
		''response.write strSql
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSql, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
				fnGetPGLogList = db3_rsget.getRows()
			END IF
		db3_rsget.close
	End Function
End Class

%>
