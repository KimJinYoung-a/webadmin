<%
'####################################################
' Description : 사은품
' History : 2010.09.27 한용민 생성
'####################################################

Class CGift
	public FGCode
	public FECode
	public FTotCnt
	public FCPage
	public FPSize
	public FSearchTxt
	public FSearchType
	public FGiftName
	public FBrand		
	public FDateType   
	public FSDate		
	public FEDate		
	public FGStatus
	public FGUsing
	public FGName  	
	public FGScope 		
	public FEGroupCode 
	public FGType      
	public FGRange1    
	public FGRange2    
	public FGKindCode  
	public FGKindType  
	public FGKindCnt   
	public FGKindlimit   
	public FRegdate    
	public FAdminid    
	public FGKindName	
    public FGKindImg
	public FGDelivery
	public FItemid
	public FOpenDate
	public FCloseDate
	public FOldKindName
	public FSiteScope
	public FPartnerID
	public Fimage120
	public FResultCount
	public FItemList
	public Fimage400List
	
	Private Sub Class_Initialize()
		redim  FItemList(0)
        redim Fimage400List(0)
        FTotCnt = 0
	    FCPage  = 1
	    FPSize  = 20
	    FResultCount = 0	  
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'== 사은품 리스트 가져오기
	public Function fnGetGiftList

		Dim strSqlCnt, strSql, strSearch,iDelCnt
		
		strSearch = ""
		IF FECode <> "" THEN
			strSearch = " and evt_code ="&FECode
		END IF	
		
		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN 
				strSearch = strSearch & " and gift_code = "&FSearchTxt
			ELSE
				strSearch = strSearch & " and evt_code = "&FSearchTxt
			END IF	
		END IF					
		
		IF FGiftName <> "" THEN
				strSearch = strSearch & " and gift_name like '%"&FGiftName&"%'"
		END IF	
		
		IF FBrand <> "" THEN
				strSearch = strSearch & " and makerid = '"&FBrand&"'"
		END IF	
		
		IF FSDate <> "" AND FEDate <> "" THEN
			if CStr(FDateType) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&FSDate&"', gift_startdate) >= 0 and  datediff(day,'"&FEDate&"', gift_startdate) <=0  "
			elseif CStr(FDateType) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&FSDate&"',gift_enddate) >= 0 and  datediff(day,'"&FEDate&"',gift_enddate) <=0  "
			end if
		END IF
		
		IF FGStatus <> "" THEN
			IF FGStatus = 9 THEN
				strSearch = strSearch & " and ( gift_status = "&FGStatus&" or  datediff(day,getdate(),gift_enddate)< 0 ) "
			ELSEIF FGStatus = 6 THEN	'오픈예정
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)<= 0 and datediff(day,getdate(),gift_enddate) >= 0  "
			ELSEIF FGStatus = 7 THEN	'오픈진행중
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)> 0 and  datediff(day,getdate(),gift_enddate)>=0 "		
			ELSE
				strSearch = strSearch & " and  gift_status = "&FGStatus&" AND  datediff(day,getdate(),gift_enddate)>=0  "
			END IF
		END IF	
	
		IF FGDelivery <> "" THEN
			strSearch = strSearch & " and gift_delivery = '"&FGDelivery&"'"
		END IF		

		strSqlCnt = " SELECT COUNT(gift_code) FROM [db_academy].[dbo].[tbl_gift] WHERE 1=1 "	&strSearch
		rsACADEMYget.Open strSqlCnt,dbACADEMYget 
		IF not rsACADEMYget.EOF THEN
			FTotCnt = rsACADEMYget(0)
		End IF
		rsACADEMYget.Close
		
		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&"  [gift_code], [gift_name], [gift_scope], [evt_code], [evtgroup_code], [makerid], [gift_type], [gift_range1], [gift_range2], A.[giftkind_code]"&_
					"  		, [giftkind_type], [giftkind_cnt], [giftkind_limit], [gift_startdate], [gift_enddate]"&_
					"		, [gift_status] = Case When DateDiff(day,getdate(),gift_enddate) < 0 Then 9 "&_
					 "							When A.gift_status = 7 and DateDiff(day,getdate(),gift_startdate) <= 0 Then 6 "&_	
					"							ELSE gift_status end "&_
					"		, A.[regdate], [gift_using], [adminid], B.giftkind_name "&_	
					"		, gift_cnt = Case gift_scope when 2 then (select count(itemid) from [db_academy].[dbo].[tbl_eventitem] WHERE evt_code = A.evt_code)"&_
					"									when 4 then (select count(itemid) from [db_academy].[dbo].[tbl_eventitem] WHERE evt_code = A.evt_code AND evtgroup_code = A.evtgroup_code)"&_
					"									when 5 then (select count(itemid) from [db_academy].[dbo].[tbl_giftitem] WHERE gift_code = A.gift_code)	"&_
					"									else 0 end "&_
					"		,gift_delivery, opendate, closedate "&_
					" FROM [db_academy].[dbo].[tbl_gift] AS A left outer join [db_academy].[dbo].[tbl_giftkind] AS B ON A.giftkind_code = B.giftkind_code "&_
					" WHERE gift_code <= ( SELECT Min(gift_code) FROM ( SELECT TOP "&iDelCnt&" gift_code "&_
					" FROM [db_academy].[dbo].[tbl_gift] WHERE 1=1 "&strSearch&" ORDER BY gift_code DESC ) as T ) "&strSearch&" ORDER BY gift_code DESC "					
			
			'response.write 	strSql
			rsACADEMYget.Open strSql,dbACADEMYget 
			IF not rsACADEMYget.EOF THEN
				fnGetGiftList = rsACADEMYget.getRows()
			End IF
			rsACADEMYget.Close
		END IF					
	End Function		

	'== 사은품 내용 보기 '/academy/gift/giftmod.asp
	public Function fnGetGiftConts
		Dim strSql
		
		strSql ="   SELECT  [gift_code], [gift_name], [gift_scope], [evt_code], [makerid], [gift_type], [gift_range1], [gift_range2], A.[giftkind_code]"&_
				"  		, [giftkind_type], [giftkind_cnt], [giftkind_limit], [gift_startdate], [gift_enddate], [gift_status], A.[regdate], [gift_using], [adminid]"&_
				"		, B.giftkind_name, B.giftkind_img, gift_delivery, opendate, closedate,lastupdate, A.gift_itemname, A.site_scope, A.partner_id "&_					
				" FROM [db_academy].[dbo].[tbl_gift] AS A left outer join [db_academy].[dbo].[tbl_giftkind] AS B ON A.giftkind_code = B.giftkind_code "&_
				" WHERE  gift_code = "&FGCode
				
		'response.write strSql &"<br>"		
		rsacademyget.Open strSql,dbacademyget 
		
		IF not rsacademyget.EOF THEN
			
			FGName  	= rsacademyget("gift_name")
			FGScope 	= rsacademyget("gift_scope")
			FECode  	= rsacademyget("evt_code")			
			FBrand      = rsacademyget("makerid")
			FGType      = rsacademyget("gift_type")
			FGRange1    = rsacademyget("gift_range1")
			FGRange2    = rsacademyget("gift_range2")
			FGKindCode  = rsacademyget("giftkind_code")
			FGKindType  = rsacademyget("giftkind_type")
			FGKindCnt   = rsacademyget("giftkind_cnt")
			FGKindlimit = rsacademyget("giftkind_limit")
			FSDate   	= rsacademyget("gift_startdate")
			FEDate     	= rsacademyget("gift_enddate")
			FGStatus    = rsacademyget("gift_status")
			FGUsing     = rsacademyget("gift_using")
			IF datediff("d",FEDate,now) > 0  THEN FGStatus = 9	'종료일이 지난 경우 종료로 표기
			'IF (datediff("d",FEDate,now) <= 0 and datediff("d",FSDate,now)>=0  and FGStatus=7) THEN FGStatus = 6	
			FRegdate    = rsacademyget("regdate")
			FAdminid    = rsacademyget("adminid")
			FGKindName	= rsacademyget("giftkind_name")
			FGKindImg	= rsacademyget("giftkind_img")
			FGDelivery  = rsacademyget("gift_delivery")
			FOpenDate	= rsacademyget("opendate")
			FCloseDate	= rsacademyget("closedate")
			FOldKindName= rsacademyget("gift_itemname")			
			FSiteScope 	= rsacademyget("site_scope")
			FPartnerID	= rsacademyget("partner_id")
			
		END IF
		rsacademyget.close
	End Function

	'== 사은품 종류 검색하기
	public Function fnGetGiftKind
		Dim strSql
		strSql = " SELECT giftkind_code, giftkind_name, giftkind_img, itemid, regdate FROM  [db_academy].[dbo].[tbl_giftkind] "&_
				" WHERE  giftkind_name like '%"&FSearchTxt&"%' order by giftkind_code desc"
		rsacademyget.Open strSql,dbacademyget 
		IF not rsacademyget.EOF THEN
			fnGetGiftKind = rsacademyget.getRows()
		End IF
		rsacademyget.Close		
	End Function	
End Class	
%>