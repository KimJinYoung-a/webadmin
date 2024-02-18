	<%
	class CComment
	public FboardIdx
	
	'//ÄÚ¸àÆ®
	public Function fnGetCommentList
	Dim strSql	 
	IF FboardIdx = "" THEN FboardIdx = 0
		strSql ="[db_board].[dbo].[sp_Ten_partnerA_notice_comment_getData]( "&FboardIdx&")"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetCommentList =  rsget.getRows()
		END IF
		rsget.close 
	End Function
 End Class
	%>