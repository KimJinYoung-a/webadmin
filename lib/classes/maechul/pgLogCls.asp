<%
'###########################################################
' Description : 결제로그_PG사별 클래스
' Hieditor : 2013.12.27 정윤정 생성
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
