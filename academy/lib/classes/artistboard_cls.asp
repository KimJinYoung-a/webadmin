<%
Class CArtistRoomBoard
    public Fidx
    
    public FCPage
    public FPSize
    public FTotCnt
    
    public FThread
    public FDepth
    public FLecuserid
    public FUserid
    public FTitle
    public FContent
    public FHit
    public FRegdate
    public FImgUrl1
    public FImgUrl2
	
    public FSearchType
    public FSearchTxt

	public Function fnGetList
		Dim strSql
		IF FTotCnt < 0 THEN
			strSql ="[db_academy].[dbo].[admin_ArtistRoom_boardCount] ('"&FLecuserid&"',"&FSearchType&",'"&FSearchTxt&"')"												
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not rsACADEMYget.EOF THEN			
				FTotCnt = 	rsACADEMYget(0)
			END IF
			rsACADEMYget.close
		END IF
		IF FtotCnt	> 0 THEN
			strSql ="[db_academy].[dbo].[admin_ArtistRoom_boardList] ('"&FLecuserid&"',"&FCPage&","&FPSize&","&FSearchType&",'"&FSearchTxt&"')"												
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not rsACADEMYget.EOF THEN
				fnGetList = rsACADEMYget.GetRows()				
			END IF				
			rsACADEMYget.close			
		END IF	
    End Function

	public Function fnGetContent
	Dim strSql
		strSql = "[db_academy].[dbo].[academy_ArtistRoom_boardView]("&Fidx&",'"&FLecuserid&"')"
		rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not rsACADEMYget.EOF THEN
				FUserid  = rsACADEMYget("userid")
				FTitle   = rsACADEMYget("title") 
				FContent  = rsACADEMYget("content")
				FImgUrl1  = rsACADEMYget("imgurl1")
				FImgUrl2  = rsACADEMYget("imgurl2")
				FHit      = rsACADEMYget("hit")
				FRegdate  = rsACADEMYget("regdate")
				FThread   = rsACADEMYget("thread")
				FDepth   = rsACADEMYget("depth")
			End IF	
		rsACADEMYget.close		
	End Function
	
	
	Private Sub Class_Initialize()
		
	End Sub
	
	Private Sub Class_Terminate()

	End Sub
end Class
%>