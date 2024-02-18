<%
Class CscmBG
	public FBGImg
	
	public Function fnGetBGUrl
		dim strSql
		strSql = " select top 1  imgUrl from db_sitemaster.dbo.tbl_scm_loginBackImg where isusing =1 order by idx desc "
		rsget.Open strSql,dbget,1
		if not rsget.eof then
			FBGImg = rsget("imgUrl") 
		end if
		rsget.close
	End Function

End Class
%>