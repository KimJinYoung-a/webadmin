<%
'####################################################
' Page : /lib/classes/event/eventSiteCls.asp
' Description :  이벤트 HTML 사이트 관리
' History : 2007.03.27 정윤정 생성
'####################################################

Class ClsEvtSite
 
 public FCPage	'Set 현재 페이지
 public FPSize	'Set 페이지 사이즈
 public FTotCnt	'Get 전체 레코드 갯수
 
 public FSIdx
 public FSLocation
 public FSType
 public FSCont 
 public FSLType
 public FSLink 
 public FSW
 public FSH
 public FSDo
 public FSUsing
				
 	public Function fnGetList
 		Dim strSql, strSqlCnt, strAdd
 		
 		IF FSLocation <> "" THEN
 			strAdd = " AND evtsite_location ="&FSLocation
 		END IF	
 		strSqlCnt = " SELECT COUNT(evtsite_idx) FROM  [db_event].[dbo].[tbl_event_sitemanage] WHERE evtsite_location > 21 "&strAdd
 		rsget.Open strSqlCnt,dbget,1
		IF not rsget.EOF THEN
			FTotCnt = rsget(0) 
		End IF
		rsget.Close	
		
		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1	
 			strSql = " SELECT TOP "&FPSize&" evtsite_idx, evtsite_location, evtsite_type, evtsite_cont, evtsite_linktype,evtsite_link,evtsite_width, evtsite_height, evtsite_disporder,evtsite_regdate, evtsite_using "&_
					"	FROM [db_event].[dbo].[tbl_event_sitemanage]"&_
					" WHERE  evtsite_location > 21 and evtsite_idx<=  ( SELECT MIN(evtsite_idx) FROM  (SELECT Top "&iDelCnt&" evtsite_idx FROM  [db_event].[dbo].[tbl_event_sitemanage] WHERE  evtsite_location > 21 "&strAdd&" ORDER BY evtsite_idx DESC ) as T ) "&_
					strAdd&" ORDER BY evtsite_idx DESC "
			rsget.Open strSql,dbget,1	
			IF not rsget.EOF THEN	
				fnGetList = rsget.getRows()
			End IF
			rsget.Close			
 		END IF
 		
 	End Function
 	
 	public Function fnGetContent
 		IF FSIdx = "" THEN Exit Function
 			
 		Dim strSql
 		strSql = "SELECT evtsite_idx, evtsite_location, evtsite_type, evtsite_cont, evtsite_linktype,evtsite_link,evtsite_width, evtsite_height, evtsite_disporder,evtsite_regdate, evtsite_using "&_
 				" FROM  [db_event].[dbo].[tbl_event_sitemanage] "&_
 				"	WHERE evtsite_idx = "&FSIdx
 		rsget.Open strSql,dbget,1	
		IF not rsget.EOF THEN	
			FSLocation = rsget("evtsite_location")
			FSType = rsget("evtsite_type")
			FSCont = rsget("evtsite_cont")
			FSLType = rsget("evtsite_linktype")
			FSLink = db2html(rsget("evtsite_link"))
			FSW= rsget("evtsite_width")
			FSH = rsget("evtsite_height")
			FSDo = rsget("evtsite_disporder")
			FSUsing = rsget("evtsite_using")
		End IF
		rsget.Close			
	End Function
End Class
%>