<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'#############################
'endtype:1 - ����������(�������ó��)
'#############################

dim sMode,itemid 
dim itemidarr,arrItem,i
dim sReturnURL
dim strSql, adminid

sMode = requestCheckvar(request("hidM"),1)
itemidarr = ReplaceRequestSpecialChar(request("itemidarr"))
sReturnURL = ReplaceRequestSpecialChar(request("sRU")) 
adminid = session("ssBctID")

SELECT CASE sMode
CASE "I"

	arrItem= split(itemidarr,",")
		
		For i=0 To ubound(arrItem)
		dbget.beginTrans
			strSql = " update [db_item].[dbo].tbl_item "&vbcrlf
			strSql = strSql & " set isusing = 'N' "&vbcrlf
			strSql = strSql & "   , sellyn = 'N'" &vbcrlf
			strSql = strSql & "   , lastupdate =getdate()" &vbcrlf
			strSql = strSql & " where " &vbcrlf
			strSql = strSql & " 	  isusing = 'Y' " &vbcrlf
			strSql = strSql & " 	and itemid  ="&trim(arrItem(i)) 
			dbget.execute strSql 
		
			If (Err) then
        response.write Err.Description
        dbget.RollBackTrans				'�ѹ�(�����߻���)
        dbget.Close
        Call Alert_msg("ó���� ������ �߻��߽��ϴ�. ")
        response.end
   		end if
   
			strSql = " insert into db_item.dbo.tbl_item_endlog "&vbcrlf
			strSql = strSql & "(itemid, endtype, adminid)"&vbcrlf
			strSql = strSql & " values "&vbcrlf
			strSql = strSql & "("&trim(arrItem(i))&" ,1, '"&adminid&"')"
			dbget.execute strSql 
			 If Err.Number = 0 Then
        	dbget.CommitTrans			
       Else
            response.write Err.Description
            dbget.RollBackTrans				'�ѹ�(�����߻���)
            dbget.Close
           Call Alert_msg("ó���� ������ �߻��߽��ϴ�." )
            response.end
        End If 	
		Next
		dbget.Close
		
		strSql = " update db_analyze_data_raw.dbo.tbl_item "&vbcrlf
		strSql = strSql & " set isusing = 'N' "&vbcrlf
		strSql = strSql & "   , sellyn = 'N'" &vbcrlf
		strSql = strSql & "   , lastupdate =getdate()" &vbcrlf
		strSql = strSql & " where " &vbcrlf
		strSql = strSql & " 	  isusing = 'Y' " &vbcrlf
		strSql = strSql & " 	and itemid in ("&itemidarr & ") " 
		dbAnalget.execute strSql 
	    dbAnalget.close
	    
		Call Alert_move("ó���Ǿ����ϴ�.", sReturnURL) 
		response.end
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT		
%>