<%
'############################
'ERP 연동 공통 데이터 가져오기
'############################

'카드사 정보 
Class CCardCorp 
'카드사 정보 가져오기
public Function fnGetCardCorp
	Dim strSql 
	IF (application("Svr_Info")="Dev") THEN
    strSql ="db_SCM_LINK.dbo.sp_CATS_CARDCORP_GETLIST_TEST"
  ELSE
    strSql ="db_SCM_LINK.dbo.sp_CATS_CARDCORP_GETLIST"
  END IF 
		dbiTms_rsget.Open strSql, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			fnGetCardCorp = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close 
End Function   
END Class
%>