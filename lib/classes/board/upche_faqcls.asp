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

'faq ���� ����Ʈ
Function fnDispFaqType(sValue)
dim strFaq
	if sValue ="1" then
		   strFaq = "��ǰ"
	elseif sValue ="2" then
		   strFaq = "�ֹ�"	   
	elseif sValue ="3" then
			strFaq = "����"	   
	elseif sValue ="4" then
			strFaq = "��Ʈ�ʰ���"	  
	elseif sValue ="5" then
			strFaq = "�����ֹ���"	  
	elseif sValue ="6" then
			strFaq = "������"	  
	else
			strFaq = "��Ÿ"	  
	end if	   
	
	fnDispFaqType = strFaq
End Function


'faq ���� selectbox
Sub sbOptFaqType(sValue)
%>
		<option value="">--����--</option>
		<option value="1" <%if sValue="1" then%>selected<%end if%>>��ǰ</option>
		<option value="2" <%if sValue="2" then%>selected<%end if%>>�ֹ�</option>
		<option value="3" <%if sValue="3" then%>selected<%end if%>>����</option>
		<option value="4" <%if sValue="4" then%>selected<%end if%>>��Ʈ�ʰ���</option>
		<option value="5" <%if sValue="5" then%>selected<%end if%>>�����ֹ���</option>
		<option value="6" <%if sValue="6" then%>selected<%end if%>>������</option>
		<option value="7" <%if sValue="7" then%>selected<%end if%>>��Ÿ</option>
 
<%
End Sub
	%>