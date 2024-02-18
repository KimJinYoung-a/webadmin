<%
 Class ClsScmMng
 
 public FTotCnt
 public FPSize
 public FCPage
 public FImgUrl 
 public Fuserid 
 public Fregdate 
 public Fusername
 public FRectIdx
 
 		public Function fnGetScmMngList
 		 dim strSql
 		 
 		 strSql = " select count(idx) FROM [db_sitemaster].[dbo].[tbl_scm_loginBackImg]  where isusing = 1 "  
 		 rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			End IF
			rsget.Close
	
 		 IF FTotCnt >0 THEN
 		 	dim iSPageNo, iEPageNo
			iSPageNo = (FPSize*(FCPage-1)) + 1
			iEPageNo = FPSize*FCPage	
		
 		 strSql = "SELECT idx, imgurl, userid, regdate  "&vbCrlf
 		 strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten where userid = TB.userid and userid <> '' order by regdate desc) as username "
 		 strSql = strSql & " FROM ( "&vbCrlf
 		 strSql = strSql &"		SELECT ROW_NUMBER() OVER (ORDER BY  idx desc ) as RowNum , " &vbCrlf
 		 strSql = strSql &"  	 idx, imgUrl, userid, regdate   "&vbCrlf 
 		 strSql = strSql & " 	FROM [db_sitemaster].[dbo].[tbl_scm_loginBackImg]  "&vbCrlf
 		 strSql = strSql & " 	WHERE isusing = 1"
 		 strSql = strSql & ") AS TB "&vbCrlf
		 strSql = strSql &	" WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " "&vbCrlf
 		 strSql = strSql & " Order by idx desc " 	
 		 rsget.Open strSql,dbget,1 
		IF not rsget.EOF THEN
			fnGetScmMngList = rsget.getRows()
		End IF
		rsget.Close
	End IF
 
 	  End Function 
 	  
 	  public Function fnGetScmMngConts
 	   dim strSql
 	   strSql = "SELECT imgurl, userid, regdate "&vbCrlf
 	   	 strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten where userid = TB.userid and userid <> '' order by regdate desc) as username "
 	    strSql = strSql & " FROM [db_sitemaster].[dbo].[tbl_scm_loginBackImg] as TB "&vbCrlf 
 		 strSql = strSql & " where idx = "&FRectIdx
 		  rsget.Open strSql,dbget,1 
		IF not rsget.EOF THEN
			FImgUrl = rsget("imgurl")
			Fuserid = rsget("userid")
			Fregdate = rsget("regdate")
			Fusername= rsget("username")
		End IF
		rsget.Close
 		End Function
 End Class
%>