<%
'############################
'ERP ���� ���� ������ ��������
'############################

'ī��� ���� 
Class CCardCorp 
'ī��� ���� ��������
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