<%
'###########################################################
' Description : ��ü reference
' Hieditor : 2016.08.19 ������ ���� 
'###########################################################


Class CRefer 

public FRectType
public FRectTitle
public FRectConts
public FRectUserName
public FRectSearch

public FTotCnt
public FPSize
public FCPage

public FrefIdx
public FTitle
public FContents
public FrefType   
public Fregid    
public Fregname  
public Fregdate  



	public Function fnGetReferList
		dim strSql, strSqlCnt, strSqlAdd
		strSqlAdd = ""
		
		if FRectType <> "" then
			strSqlAdd  = strSqlAdd & " and reftype = "&FRectType
		end if
		
		if FRectSearch <> "" then
			strSqlAdd  = strSqlAdd & " and (Title like '%"&FRectSearch&"%' or contents like  '%"&FRectSearch&"%')"
		end if
		 
		
		strSqlCnt = " SELECT count(refidx) "&vbcrlf
		strSqlCnt = strSqlCnt & " from  db_board.dbo.tbl_partnerA_reference as f "&vbcrlf
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
		strSql = " SELECT refidx, reftype, title, contents, regid, regdate, username  "&vbcrlf
		strSql = strSql &" FROM ( "&vbcrlf
		strSql = strSql &" 		SELECT  ROW_NUMBER() OVER (ORDER BY  refidx desc ) as RowNum ,f.refidx, f.reftype, f.title, f.contents, f.regid, f.regdate, t.username "&vbcrlf 
		strSql = strSql & " 	from  db_board.dbo.tbl_partnerA_reference as f "&vbcrlf
		strSql = strSql & "			inner join db_partner.dbo.tbl_user_tenbyten as t on f.regid = t.userid "&vbcrlf
		strSql = strSql & " 	where f.isusing =1  " & strSqlAdd&vbcrlf
		strSql = strSql &") AS TB "&VbCrlf 		
		strSql = strSql &" WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " "&VbCrlf		
		strSql = strSql &" order by  TB.RowNum  " 
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				fnGetReferList = rsget.getRows()
			End IF
		rsget.Close
				
		END IF
	End Function
	
	public Function fnGetReferConts
		Dim strSql
		strSql =" SELECT f.refidx, f.reftype, f.title, f.contents, f.regid, f.regdate, t.username "&vbcrlf
		strSql = strSql & " 	from  db_board.dbo.tbl_partnerA_reference as f "&vbcrlf
		strSql = strSql & "			inner join db_partner.dbo.tbl_user_tenbyten as t on f.regid = t.userid "&vbcrlf
		strSql = strSql & " 	where f.refidx = " &FrefIdx 
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				FrefIdx   =rsget("refidx")         
				FrefType  =rsget("reftype")        
				FTitle 		=rsget("title")          
				FContents =rsget("contents")      
				Fregid    =rsget("regid")        
				Fregname  =rsget("username") 
				Fregdate  =rsget("regdate")             
			END IF
		rsget.Close	
		
	End Function
  
	public Function fnGetAttachFile
		dim strSql
		strSql = "SELECT attachFileidx,refIdx,fileLink FROM db_board.dbo.tbl_partnerA_reference_attachfile WHERE refidx ="&FrefIdx
		 
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				 fnGetAttachFile = rsget.getRows()
			END IF
		rsget.close
	End Function
End Class

'reference ���� selectbox
Function sbOptReferType(sValue)
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
End Function

'reference ���� ����Ʈ
Function fnDispReferType(sValue)
dim strref
	if sValue ="1" then
		   strref = "��ǰ"
	elseif sValue ="2" then
		   strref = "�ֹ�"	   
	elseif sValue ="3" then
			strref = "����"	   
	elseif sValue ="4" then
			strref = "��Ʈ�ʰ���"	  
	elseif sValue ="5" then
			strref = "�����ֹ���"	  
	elseif sValue ="6" then
			strref = "������"	  
	else
			strref = "��Ÿ"	  
	end if	   
	
	fnDispReferType = strref
End Function
%>
