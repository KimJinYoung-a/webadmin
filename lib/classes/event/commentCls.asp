	<%
	class CComment
	public FEvtCode
	
	'//ÄÚ¸àÆ®
	public Function fnGetCommentList
	Dim strSql	 
	IF FEvtCode = "" THEN FEvtCode = 0 
  
		strSql =" 	SELECT A.comidx, A.comment, A.regId, A.regdate,	case a.regtype when 'U' then C.socname_kor	else B.username end as username    "
		strSql = strSql &" FROM db_event.dbo.tbl_partner_event_comment  AS A "
		strSql = strSql &" left outer JOIN db_partner.dbo.tbl_user_tenbyten AS B ON A.regId = B.userid and a.regtype='A'"
		strSql = strSql &" left outer  JOIN db_user.dbo.tbl_user_c as C on A.regid = C.userid and a.regtype ='U'"
		strSql = strSql &" WHERE A.evt_code = "&FEvtCode&"  and A.isUsing = 1 "
		strSql = strSql &" Order by A.comidx "	  
			rsget.Open strSql,dbget,1
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			fnGetCommentList =  rsget.getRows()
		END IF
		rsget.close 
	End Function
 End Class
	%>