<%
Class CFaq

public FRectType
public FRectTitle
public FRectConts
public FRectUserName

public FTotCnt
public FPSize
public FCPage

public FFaqIdx
public FTitle
public FContents
public FFaqType   
public Fregid    
public Fregname  
public Fregdate  
public FRectSearch


	public Function fnGetFaqList
		dim strSql, strSqlCnt, strSqlAdd
		strSqlAdd = ""
		
		if FRectType <> "" then
			strSqlAdd  = strSqlAdd & " and faqtype = "&FRectType
		end if
		
		if FRectSearch <> "" then
			strSqlAdd  = strSqlAdd & " and ( Title like '%"&FRectSearch&"%' or contents  like '%"&FRectSearch&"%' )"
		end if
		 
		strSqlCnt = " SELECT count(faqidx) "&vbcrlf
		strSqlCnt = strSqlCnt & " from  db_board.dbo.tbl_partnerA_faq as f "&vbcrlf
		strSqlCnt = strSqlCnt & "		inner join db_partner.dbo.tbl_user_tenbyten as t on f.regid = t.userid "&vbcrlf
		strSqlCnt = strSqlCnt & " where f.isusing =1  " & strSqlAdd
		rsget.Open strSqlCnt,dbget 
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close
		
		IF FTotCnt >0 THEN
				dim iSPageNo, iEPageNo
				iSPageNo = (FPSize*(FCPage-1)) + 1
				iEPageNo = FPSize*FCPage	
		strSql = " SELECT faqidx, faqtype, title, contents, regid, regdate, username "&vbcrlf
		strSql = strSql &" FROM ( "&vbcrlf
		strSql = strSql &" 		SELECT  ROW_NUMBER() OVER (ORDER BY  faqidx desc ) as RowNum ,f.faqidx, f.faqtype, f.title, f.contents, f.regid, f.regdate, t.username "&vbcrlf
		strSql = strSql & " 	from  db_board.dbo.tbl_partnerA_faq as f "&vbcrlf
		strSql = strSql & "			inner join db_partner.dbo.tbl_user_tenbyten as t on f.regid = t.userid "&vbcrlf
		strSql = strSql & " 	where f.isusing =1  " & strSqlAdd&vbcrlf
		strSql = strSql &") AS TB "&VbCrlf 		
		strSql = strSql &" WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " "&VbCrlf		
		strSql = strSql &" order by  TB.RowNum  " 
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				fnGetFaqList = rsget.getRows()
			End IF
		rsget.Close
				
		END IF
	End Function
	
		public Function fnGetFaqConts
		Dim strSql
		strSql =" SELECT f.faqidx, f.faqtype, f.title, f.contents, f.regid, f.regdate, t.username "&vbcrlf
		strSql = strSql & " 	from  db_board.dbo.tbl_partnerA_faq as f "&vbcrlf
		strSql = strSql & "			inner join db_partner.dbo.tbl_user_tenbyten as t on f.regid = t.userid "&vbcrlf
		strSql = strSql & " 	where f.faqidx = " &FFaqIdx 
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				FFaqIdx   =rsget("faqidx")         
				FFaqType  =rsget("faqtype")        
				FTitle 		=rsget("title")          
				FContents =rsget("contents")      
				Fregid    =rsget("regid")        
				Fregname  =rsget("username") 
				Fregdate  =rsget("regdate")             
			END IF
		rsget.Close	
		
	End Function
End Class
'--=========================================================================================

'faq 구분 리스트
Function fnDispFaqType(sValue)
dim strFaq
	if sValue ="1" then
		   strFaq = "상품"
	elseif sValue ="2" then
		   strFaq = "주문"	   
	elseif sValue ="3" then
			strFaq = "정산"	   
	elseif sValue ="4" then
			strFaq = "파트너관리"	  
	elseif sValue ="5" then
			strFaq = "물류주문서"	  
	elseif sValue ="6" then
			strFaq = "오프샵"	  
	else
			strFaq = "기타"	  
	end if	   
	
	fnDispFaqType = strFaq
End Function


'faq 구분 selectbox
Sub sbOptFaqType(sValue)
%>
		<option value="">--선택--</option>
		<option value="1" <%if sValue="1" then%>selected<%end if%>>상품</option>
		<option value="2" <%if sValue="2" then%>selected<%end if%>>주문</option>
		<option value="3" <%if sValue="3" then%>selected<%end if%>>정산</option>
		<option value="4" <%if sValue="4" then%>selected<%end if%>>파트너관리</option>
		<option value="5" <%if sValue="5" then%>selected<%end if%>>물류주문서</option>
		<option value="6" <%if sValue="6" then%>selected<%end if%>>오프샵</option>
		<option value="7" <%if sValue="7" then%>selected<%end if%>>기타</option>
 
<%
End Sub
	%>