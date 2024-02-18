<%
Class CWonderday
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
	
	public FIdx
	'// 리스트
	public Function fnGetImgList
		Dim strSql,iDelCnt
		strSql = " SELECT COUNT(idx) FROM [db_event].[dbo].[tbl_event_wonderday] "
		rsget.Open strSql,dbget 
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF	
		rsget.close
		
		IF FTotCnt > 0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1	
			
			strSql = " SELECT  TOP "&FPSize&"  idx, listimg, isUsing, regdate, opendate, volnum "&_
					" FROM [db_event].[dbo].[tbl_event_wonderday] "&_
					" WHERE idx <= (SELECT min(idx) FROM ( SELECT TOP "&iDelCnt&" idx FROM [db_event].[dbo].[tbl_event_wonderday]  ORDER BY idx DESC ) as T ) "&_
					" ORDER BY idx DESC "					
			rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				fnGetImgList = rsget.getRows() 
			END IF	
			rsget.close
		END IF	
	End Function
	
	public FListImg
	public FMainImg
	public FUsing
	public FRegdate
	public FOpendate
	public FVolnum
	
	'// 내용보기
	public Function fnGetConts
	 	Dim strSql
	 	strSql = " SELECT idx, listimg, mainimg, isUsing, regdate, opendate, volnum "&_
	 			" FROM [db_event].[dbo].[tbl_event_wonderday] "&_
				" WHERE idx ="&FIdx				
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				FListImg	= rsget("listimg")	
				FMainImg	= rsget("mainimg")	
				FUsing		= rsget("isUsing")	
				FRegdate	= rsget("regdate")	
				FOpendate	= rsget("opendate")
				FVolnum		= rsget("volnum")
			END IF	
		rsget.close				
	END Function			
End Class
%>
