<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �����ɽ� �̺�Ʈ db���
' History : 2008.05.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
 Dim sMode,  blnUse, sListImg, sMainImg, iIdx,iCurrentpage, dOpendate, iVolnum
 Dim strSql
 
 sMode 		=  requestCheckVar(Request.Form("sM"),1)
 menupos 	=  requestCheckVar(Request.Form("menupos"),10)
 blnUse 	=  requestCheckVar(Request.Form("blnU"),1)
 sListImg 	=  requestCheckVar(Request.Form("sLImg"),100)
 sMainImg 	=  requestCheckVar(Request.Form("sMImg"),100)
 dOpendate  =  requestCheckVar(Request.Form("dOD"),10)
 iIdx 		=  requestCheckVar(Request.Form("idx"),10)
 iCurrentpage =  requestCheckVar(Request.Form("iC"),10)
 iVolnum	=	requestCheckVar(Request.Form("iVN"),3)
 IF blnUse = "" THEN blnUse = 0
 	 
 SELECT Case sMode
 	Case "I"
 		strSql = " INSERT INTO [db_event].[dbo].tbl_event_wonderday (listImg, mainImg, isUsing, adminID, opendate,volnum)" &_
 				"	VALUES ('"&sListImg&"','"&sMainImg&"', "&blnUse&", '"&session("ssBctId")&"','"&dOpendate&"',"&iVolnum&")" 			
 		dbget.execute strSql
 		
 		IF Err.Number = 0 THEN	
			Call sbAlertMsg ("��ϵǾ����ϴ�.", "index.asp?menupos="&menupos, "self") 		
		ELSE		
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�", "back", "") 
		END IF
		dbget.close()	:	response.End		
 	Case "U" 	
	 	strSql = " UPDATE [db_event].[dbo].tbl_event_wonderday "&_
	 			 " SET listImg = '"&sListImg&"', mainImg='"&sMainImg&"', isUsing= "&blnUse&",  adminID = '"&session("ssBctId")&"' , opendate = '"&dOpendate&"', volnum="&iVolnum&_
	 			 " WHERE idx = "&iIdx
	 		dbget.execute strSql	 	
	 		IF Err.Number = 0 THEN	
				Call sbAlertMsg ("�����Ǿ����ϴ�.", "index.asp?menupos="&menupos&"&iC="&iCurrentpage, "self") 		
			ELSE		
				Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�", "back", "") 
			END IF
			dbget.close()	:	response.End
 	Case ELSE
 		Call sbAlertMsg ("�Ķ���Ͱ��� ������ �߻��Ͽ����ϴ�.", "back", "") 
 END SELECT	
%>
