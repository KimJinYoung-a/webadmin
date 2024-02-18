<%@ language=vbscript %>
<% option explicit %> 
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   
 
'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  �̺�Ʈ ���� ������ó�� - ���, ����, ����
' History : 2007.02.12 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<%
Dim smode
Dim eCode,ePCode,evtgroup_code,eGDepth,eGDesc,eGSort,eChannel, sTarget,eModeType
dim strSql,strMsg
dim eGCodeArr, ePGCodeArr,eSortArr, newPCode, newDepth,intLoop
Dim eGIsDisp, vChangeContents, vSCMChangeSQL, eGbrand, etype, linkkind
  
smode = requestCheckVar(Request("mode"),2)
eCode = requestCheckVar(Request("eC"),10)
eChannel= requestCheckVar(Request("eCh"),1)
evtgroup_code = requestCheckVar(Request("eGC"),10)
ePCode= requestCheckVar(Request("selPC"),10)
eGDesc= requestCheckVar(Request("sGD"),32)
eGSort= requestCheckVar(Request("sGS"),10)
sTarget= requestCheckVar(Request("sTarget"),10)
eModeType= requestCheckVar(Request("eMT"),1)
eGIsDisp= requestCheckVar(Request("eIsDisp"),1)
eGbrand= requestCheckVar(Request("eGbrand"),32)
etype= requestCheckVar(Request("etype"),2)
linkkind = requestCheckVar(Request("linkkind"),1)

SELECT CASE smode 
CASE "I" 
		IF ePCode = "0" THEN
		 	strSql = "select isnull(max(evtgroup_depth),0) + 100 FROM  [db_event].[dbo].[tbl_eventitem_group] where evt_code = "&eCode 
		ELSE	
			strSql = "select isnull(max(evtgroup_depth),0)+1 FROM  [db_event].[dbo].[tbl_eventitem_group] WHERE evt_code = "&eCode&" and (evtgroup_code = "& ePCode&" OR evtgroup_pcode ="&ePCode&")  "
		END IF
	
			rsget.Open strSql, dbget
			IF not (rsget.EOF or rsget.BOF) THEN
				eGDepth = 	rsget(0)
			END IF	
			rsget.Close
			
			strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'"&eGDesc& "', "&eGSort&","&ePCode&","&eGDepth&",'"&eGDesc& "', "&eGSort&","&ePCode&","&eGDepth&")"	 
    		dbget.execute strSql	
    		
    		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_eventitem_group] "	'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
    		strSql = "select SCOPE_IDENTITY()"

    		rsget.Open strSql, dbget, 0
    		evtgroup_code = rsget(0)
    		rsget.Close
		 	
		 	strSql = "UPDATE db_event.dbo.tbl_eventitem_Group set evtgroup_code_mo = evtgroup_code"
			If eChannel="P" Then
			strSql = strSql & " ,evtgroup_brand='" & eGbrand & "'"
			strSql = strSql & " ,evtgroup_linkkind='" & linkkind & "'"
			strSql = strSql & " ,evtgroup_brand_mo='" & eGbrand & "'"
			strSql = strSql & " ,evtgroup_linkkind_mo='" & linkkind & "'"
			Else
			strSql = strSql & " ,evtgroup_brand_mo='" &eGbrand& "'"
			strSql = strSql & " ,evtgroup_linkkind_mo='" & linkkind & "'"
			End If
			strSql = strSql & " where evt_code = "&eCode&" and evtgroup_code = "&evtgroup_code
	        dbget.execute strSql
	        
		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & evtgroup_code & vbCrLf
		vChangeContents = vChangeContents & "- ���� �׷� = " & ePCode & vbCrLf
		vChangeContents = vChangeContents & "- �׷�� = " & eGDesc & vbCrLf
		vChangeContents = vChangeContents & "- ���ļ��� = " & eGSort & vbCrLf
		vChangeContents = vChangeContents & "- ���ÿ��� = " & eGIsDisp & vbCrLf
    	'### ���� �α� ����(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
	        
		Call sbAlertMsg ("��ϵǾ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")
    	response.end
CASE "U" 

     if eChannel = "P" then
		strSql = "UPDATE  [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_desc ='"&eGDesc&"',evtgroup_desc_mo='"&eGDesc&"'"&_
					", evtgroup_sort='"&eGSort&"', evtgroup_sort_mo='"&eGSort&"', evtgroup_pcode = "&ePCode&", evtgroup_pcode_mo = "&ePCode&""&_
					" , evtgroup_isDisp ="&eGIsDisp&", evtgroup_isDisp_mo ="&eGIsDisp&" , evtgroup_brand='"&eGbrand&"', evtgroup_brand_mo='"&eGbrand&"'"&_
					" , evtgroup_linkkind='"&linkkind&"', evtgroup_linkkind_mo='"&linkkind&"'"&_
					" WHERE evtgroup_code ="&evtgroup_code
	 else
	    strSql = "UPDATE  [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_desc_mo='"&eGDesc&"', evtgroup_sort_mo='"&eGSort&"' , evtgroup_linkkind_mo='"&linkkind & "'"&_
					"	,evtgroup_pcode_mo="&ePCode&" , evtgroup_isDisp_mo="&eGIsDisp&" , evtgroup_brand_mo='"&eGbrand&"'"&_
					" WHERE evtgroup_code ="&evtgroup_code
	 end If
	 'Response.write strSql
	 'Response.end
		dbget.execute strSql
		
		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & evtgroup_code & vbCrLf
		vChangeContents = vChangeContents & "- ���� �׷� = " & ePCode & vbCrLf
		vChangeContents = vChangeContents & "- �׷�� = " & eGDesc & vbCrLf
		vChangeContents = vChangeContents & "- ���ļ��� = " & eGSort & vbCrLf
		vChangeContents = vChangeContents & "- ���ÿ��� = " & eGIsDisp & vbCrLf
    	'### ���� �α� ����(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
		
		Call sbAlertMsg ("�����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")
	response.end
CASE "D" 
		strSql = "UPDATE  [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_using = 'N'	"&_ 
					" WHERE evtgroup_code ="&evtgroup_code	 												
		dbget.execute strSql		
		
		strSql = "delete from [db_event].[dbo].[tbl_eventitem] WHERE evtgroup_code ="&evtgroup_code
		dbget.execute strSql	

		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & evtgroup_code & vbCrLf

    	'### ���� �α� ����(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)

		Call sbAlertMsg ("�����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")	
	response.end
CASE "GS" '�׷����, �׷��  ��������
	  dim eGDescArr, eGBArr, eGLKArr
	  
	    eGCodeArr = split(request("eGCArr"),",")
	    ePGCodeArr = split(request("ePGCArr"),",")
	    eSortArr = split(request("eSArr"),",")
	    eGDescArr= split(html2db(request("sGDArr")),"|")
		eGBArr= split(html2db(request("sGBarr")),"|")
		eGLKArr= split(html2db(request("eGLKArr")),"|")
	    newDepth = 0
	    newPcode = 0
	    
	    IF isARRay(eGCodeArr) THEN
    	    For intLoop = 0 To UBound(eGCodeArr)
    	        if ePGCodeArr(intLoop) = 0 then
    	            newPcode = eGCodeArr(intLoop) 
    	            newDepth = (Cint(newDepth*0.01)*100) + 100
    	        end if
    	           
    	      if eChannel = "P" then
        	    strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_sort = "&trim(eSortArr(intLoop)) 
        	            if ePGCodeArr(intLoop) = 0 then
        	    strSql =  strSql&", evtgroup_pcode = 0"
				strSql =  strSql&", evtgroup_pcode_mo = 0"  
        	            else    
        	    strSql =  strSql&", evtgroup_pcode = "&newPcode
				strSql =  strSql&", evtgroup_pcode_mo = "&newPcode 
        	            end if    
        	    strSql =  strSql&"  , evtgroup_depth =   "&newDepth
				strSql =  strSql&"  , evtgroup_depth_mo =   "&newDepth 
        	    strSql = strSql & " , evtgroup_desc ='"&trim(eGDescArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_desc_mo ='"&trim(eGDescArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_brand='"&trim(eGBArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_brand_mo='"&trim(eGBArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_linkkind='" & trim(eGLKArr(intLoop)) & "'"
				strSql = strSql & " , evtgroup_linkkind_mo='" & trim(eGLKArr(intLoop)) & "'"
        	    strSql =  strSql&" where evtgroup_code = "&trim(eGCodeArr(intLoop))&"  and evt_code = "&eCode
    	     else
    	        strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_sort_mo = "&trim(eSortArr(intLoop)) 
        	            if ePGCodeArr(intLoop) = 0 then
        	    strSql =  strSql&", evtgroup_pcode_mo = 0"  
        	            else    
        	    strSql =  strSql&", evtgroup_pcode_mo = "&newPcode 
        	            end if    
        	    strSql =  strSql&"  , evtgroup_depth_mo =   "&newDepth 
        	    strSql = strSql & " , evtgroup_desc_mo ='"&trim(eGDescArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_brand_mo='"&trim(eGBArr(intLoop))&"'"
				strSql = strSql & " , evtgroup_linkkind_mo='" & trim(eGLKArr(intLoop)) & "'"
        	    strSql =  strSql&" where evtgroup_code_mo = "&trim(eGCodeArr(intLoop))&"  and evt_code = "&eCode
        	     
             end if 
    	 
    	    dbget.execute strSql
    	    newDepth = newDepth + 1
    	    
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & trim(eGCodeArr(intLoop)) & vbCrLf
			vChangeContents = vChangeContents & "- �׷�� = " & trim(eGDescArr(intLoop)) & vbCrLf
			vChangeContents = vChangeContents & "- ���ļ��� = " & trim(eSortArr(intLoop)) & vbCrLf
	    	'### ���� �α� ����(event)
	    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
	    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	    	dbget.execute(vSCMChangeSQL)
    	    
    	    Next
	    ELSE
	       if eChannel = "P" then
        	    strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_sort = "&eSortArr  
        	    strSql = strSql&", evtgroup_pcode = 0"   
        	    strSql = strSql&"  , evtgroup_depth =   "&newDepth 
        	    strSql = strSql & " , evtgroup_desc ='"&eGDescArr&"'"
				strSql = strSql & " , evtgroup_brand='"&trim(eGBArr)&"'"
				strSql = strSql & " , evtgroup_linkkind='" & trim(eGLKArr) & "'"
				strSql = strSql&"  , evtgroup_depth_mo =   "&newDepth 
        	    strSql = strSql & " , evtgroup_desc_mo ='"&eGDescArr&"'"
				strSql = strSql & " , evtgroup_brand_mo='"&trim(eGBArr)&"'"
				strSql = strSql & " , evtgroup_linkkind_mo='" & trim(eGLKArr) & "'"

        	    strSql = strSql&" where evtgroup_code = "&eGCodeArr&"  and evt_code = "&eCode
    	     else
    	        strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_sort_mo = "&eSortArr  
        	    strSql =  strSql&", evtgroup_pcode_mo = 0"   
        	    strSql =  strSql&"  , evtgroup_depth_mo =   "&newDepth 
        	    strSql = strSql & " , evtgroup_desc_mo ='"& eGDescArr &"'"
				strSql = strSql & " , evtgroup_brand_mo='"&trim(eGBArr)&"'"
				strSql = strSql & " , evtgroup_linkkind_mo='" & trim(eGLKArr) & "'"
        	    strSql =  strSql&" where evtgroup_code_mo = "& eGCodeArr &"  and evt_code = "&eCode
        	     
             end if
             
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & eGCodeArr & vbCrLf
			vChangeContents = vChangeContents & "- �׷�� = " & eGDescArr & vbCrLf
			vChangeContents = vChangeContents & "- ���ļ��� = " & eSortArr & vbCrLf
	    	'### ���� �α� ����(event)
	    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
	    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & eGCodeArr & "', '" & menupos & "', "
	    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	    	dbget.execute(vSCMChangeSQL)
	    	
	    End IF	
	   	'Call sbAlertMsg ("����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")
		strMsg="����Ǿ����ϴ�."
	'response.end 
CASE "J" '�׷���ġ��
         dim eGPMCode,eGMdesc,eGMCode,eGMSort, eGMDepth
    	eGCodeArr = split(request("eGCArr"),",")
    	 
    	strSql = "SELECT  evtgroup_pcode_mo , evtgroup_desc_mo, evtgroup_code_mo, evtgroup_sort_mo, evtgroup_depth_mo FROM [db_event].[dbo].[tbl_eventitem_group] WHERE evt_code = "&eCode&" and evtgroup_code =" &eGCodeArr(0)&" order by evtgroup_code asc"
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			 eGPMCode = rsget("evtgroup_pcode_mo")
			 eGMdesc = rsget("evtgroup_desc_mo")
			 eGMCode = rsget("evtgroup_code_mo")
			 eGMSort = rsget("evtgroup_sort_mo")
			 eGMDepth = rsget("evtgroup_depth_mo")
		END IF	
		
    	For intLoop = 1 To UBound(eGCodeArr) 
    	   strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] "&_
    	            " SET evtgroup_pcode_mo = "&eGPMCode&", evtgroup_desc_mo ='"&eGMdesc&"',evtgroup_sort_mo ="&eGMSort&", evtgroup_code_mo=" &eGMCode&" , evtgroup_depth_mo ="&eGMDepth&_
    	            " WHERE evt_code =  "&eCode&" and evtgroup_code ="&eGCodeArr(intLoop)  
    	  dbget.execute strSql 
        Next
        
		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ��ġ��" & vbCrLf
		vChangeContents = vChangeContents & "- �׷��ڵ� = " & eGCodeArr & vbCrLf
		'### ���� �α� ����(event)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(vSCMChangeSQL)
        
    	 	Call sbAlertMsg ("����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")	 
	response.end 
CASE "Di" '�׷쳪����
	    strSql = "  UPDATE [db_event].[dbo].[tbl_eventitem_group] "&_
	           " SET evtgroup_code_mo = evtgroup_code ,evtgroup_depth_mo=evtgroup_depth , evtgroup_pcode_mo = evtgroup_pcode "&_
	          " where evtgroup_code_mo = "&evtgroup_code& " and evt_code =" &eCode
	          
	     dbget.execute strSql 
	     
		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ������. �ڵ� = " & evtgroup_code & vbCrLf
		'### ���� �α� ����(event)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(vSCMChangeSQL)
	     
	       Call sbAlertMsg ("����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")	 
	response.end     
Case "A"	 '���ü���
    if eChannel = "P" then
        strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] set evtgroup_isDisp = "&eGIsDisp&" ,evtgroup_isDisp_mo = "&eGIsDisp&" where  evtgroup_code = "&evtgroup_code& " and evt_code =" &eCode
        dbget.execute strSql 
	else
	     strSql = " UPDATE [db_event].[dbo].[tbl_eventitem_group] set evtgroup_isDisp_mo = "&eGIsDisp&" where  evtgroup_code = "&evtgroup_code& " and evt_code =" &eCode
	     dbget.execute strSql 
    end if

	vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����. �ڵ� = " & evtgroup_code & vbCrLf
	vChangeContents = vChangeContents & "- ���ÿ��� = " & eGIsDisp & vbCrLf
	'### ���� �α� ����(event)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

	       Call sbAlertMsg ("����Ǿ����ϴ�.",  "pop_eventitem_group.asp?eC="&eCode&"&eCh="&eChannel&"&sTarget="&sTarget, "self")	 
	response.end  
    
CASE "F" '//�⺻����   
		strSql = "SELECT isNull(evtgroup_code,0) FROM [db_event].[dbo].[tbl_eventitem_group] WHERE evt_code = "&eCode &" and evtgroup_using = 'Y'"
		rsget.Open strSql, dbget
			IF not (rsget.EOF or rsget.BOF) THEN
				evtgroup_code = 	rsget(0)
			END IF	
		IF evtgroup_code <> "" THEN
				Call sbAlertMsg ("�̹� ��ϵ� �ڵ尡 �����մϴ�.", "close", "self")
			rsget.close
			response.End 
		END IF	
		rsget.close
		dim eDepth
		For intLoop =1  To eModeType
		eDepth = 100*intLoop
		
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Tab"&intLoop&"', 0,0,"&eDepth&", 'Tab"&intLoop&"',0,0,"&eDepth&")"	 
    		dbget.execute strSql		
    		
    		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_eventitem_group] "	'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
    		strSql = "select SCOPE_IDENTITY()"

    		rsget.Open strSql, dbget, 0
    		ePCode = rsget(0)
    		rsget.Close
      
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Sub"&intLoop&"_1', 1,"&ePCode&","&(eDepth+1)&",'Sub"&intLoop&"_1', 1,"&ePCode&","&(eDepth+1)&")"			
    		dbget.execute strSql		
    		
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Sub"&intLoop&"_2', 2,"&ePCode&","&(eDepth+2)&",'Sub"&intLoop&"_2', 2,"&ePCode&","&(eDepth+2)&")"			
    		dbget.execute strSql		
    		
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Sub"&intLoop&"_3',3,"&ePCode&","&(eDepth+3)&",'Sub"&intLoop&"_3', 3,"&ePCode&","&(eDepth+3)&")"			
    		dbget.execute strSql	
    		
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Sub"&intLoop&"_4',4,"&ePCode&","&(eDepth+4)&",'Sub"&intLoop&"_4', 4,"&ePCode&","&(eDepth+4)&")"			
    		dbget.execute strSql
    		
    		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode,evtgroup_depth, evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_pcode_mo,evtgroup_depth_mo) "	&_
    				" VALUES ("&eCode&",'Sub"&intLoop&"_5',5,"&ePCode&","&(eDepth+5)&",'Sub"&intLoop&"_5', 5,"&ePCode&","&(eDepth+5)&")"			
    		dbget.execute strSql 
	    Next
	    
	        '����Ͽ� ���ñ׷��ڵ�(evtgroup_code_mo)�� �ʱ⼳���� �׷��ڵ�� �����ϰ�
	        strSql = "UPDATE db_event.dbo.tbl_eventitem_Group set evtgroup_code_mo = evtgroup_code where evt_code = "&eCode
	        dbget.execute strSql 
	        
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ����." & vbCrLf
			vChangeContents = vChangeContents & "- Tab"&eModeType&" + ���� 5 �⺻�׷� ����" & vbCrLf
			'### ���� �α� ����(event)
			vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
			vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & evtgroup_code & "', '" & menupos & "', "
			vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
			dbget.execute(vSCMChangeSQL)
	        
		    strMsg = " Tab"&eModeType&" + ���� 5 �⺻�׷��� �����Ǿ����ϴ�." 
CASE "C"	 '�׷���ü����
	    strSql = "SELECT isNull(evtgroup_code,0) FROM [db_event].[dbo].[tbl_eventitem_group] WHERE evt_code = "&eCode&" and evt_channel ='"&eModeType&"'"  
		rsget.Open strSql, dbget
			IF not (rsget.EOF or rsget.BOF) THEN
				evtgroup_code = 	rsget(0)
			END IF	
		IF evtgroup_code <> "" THEN
				Call sbAlertMsg ("�̹� ��ϵ� �׷��ڵ尡 �����մϴ�.", "close", "self")
			rsget.close
			response.End 
		END IF	
		rsget.close
	
	    strSql = " insert into db_event.dbo.tbl_eventitem_group(evt_code,evtgroup_desc, evtgroup_sort, evtgroup_pcode, evtgroup_depth,evt_channel) "&vbCrlf&_
	             " ( select evt_code, evtgroup_desc, evtgroup_sort, evtgroup_pcode,  evtgroup_depth, '"&eModeType&"' "&vbCrlf&_
	             "  from db_event.dbo.tbl_eventitem_Group  "&vbCrlf&_
	             "   where evt_code = "&eCode&" and evt_channel <> '"&eModeType&"'  and evtgroup_pcode = 0  and evtgroup_using ='Y' "&vbCrlf&_
	             " )"
	      dbget.execute strSql
	    
	    strSql = " insert into db_event.dbo.tbl_eventitem_group(evt_code,evtgroup_desc, evtgroup_sort,evtgroup_pcode, evtgroup_depth,evt_channel) "&vbCrlf&_
	             " ( select evt_code, evtgroup_desc, evtgroup_sort,NULL,  evtgroup_depth, '"&eModeType&"' "&vbCrlf&_
	             "  from db_event.dbo.tbl_eventitem_Group  "&vbCrlf&_
	             "   where evt_code = "&eCode&" and evt_channel <> '"&eModeType&"' and evtgroup_pcode > 0  and evtgroup_using ='Y' "&vbCrlf&_
	             " )"
	     dbget.execute strSql
	   
	    
	    
	    strSql =  " update A "&vbCrlf&_ 
	             "       set  evtgroup_pcode =  ( select evtgroup_code from db_Event.dbo.tbl_Eventitem_Group where evt_code = A.evt_code and evt_channel = A.evt_channel and evtgroup_using ='Y'  and round(evtgroup_depth*0.01,0,1) = round(A.evtgroup_depth*0.01,0,1) and evtgroup_pcode = 0) "&vbCrlf&_ 
	             "   from    db_event.dbo.tbl_eventitem_Group  as A  "  &vbCrlf&_
	             "   where A.evt_code = "&eCode&" and A.evt_channel = '"&eModeType&"'   and A.evtgroup_using ='Y' and A.evtgroup_pcode is null  "   
	    dbget.execute strSql  
	   
	    strSql = " insert into db_event.dbo.tbl_eventitem (evt_code,itemid,evtgroup_code,evtitem_sort,evt_channel) " &vbCrlf&_
	            " ( select i.evt_code, i.itemid , gm.evtgroup_code , i.evtitem_sort, '"&eModeType&"' " &vbCrlf&_
	            "        from db_event.dbo.tbl_eventitem as i " &vbCrlf&_
	            "            left outer join db_event.dbo.tbl_eventitem_group as g " &vbCrlf&_
		        "               on i.evt_code = g.evt_code and i.evtgroup_code = g.evtgroup_code  and g.evtgroup_using = 'Y' and g.evt_channel ='"&eChannel&"' " &vbCrlf&_
                "	         left outer join db_event.dbo.tbl_eventitem_Group as gm " &vbCrlf&_
	            "                on i.evt_code = gm.evt_code and g.evtgroup_depth = gm.evtgroup_depth  and gm.evtgroup_using ='Y' and gm.evt_channel ='"&eModeType&"' " &vbCrlf&_
                "        where   i.evt_code = "&eCode&" and i.evt_channel ='"&eChannel&"'  " &vbCrlf&_
	            " )" 
	    dbget.execute strSql    

		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ��ü ����." & vbCrLf
		vChangeContents = vChangeContents & "- tbl_eventitem_Group.evt_channel = " & eModeType & vbCrLf
		'### ���� �α� ����(event)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(vSCMChangeSQL)

	    strMsg = "��ü���簡 �Ϸ�Ǿ����ϴ�."
	   
END SELECT	

	Dim arrList,intg  ,cEGroup 
        set cEGroup = new ClsEventGroup
     	cEGroup.FECode = eCode
     	cEGroup.FEChannel = eChannel  	 
      	arrList = cEGroup.fnGetEventItemGroup	 
        set cEGroup = nothing	
%>
<div id="divIpG" style="display:none;">
<%IF isArray(arrList) THEN %>
	<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
		<td>�׷��ڵ�</td>					
		<td>�����׷�</td>
		<td>�׷��</td>
		<td>���ļ���</td>
		<% If etype<>"MD" Then %>
		<td>�̹���</td>
		<% End If %>
		<td>����</td>
	</tr>
	<%FOR intg = 0 To UBound(arrList,2)%>				   						
	<tr>
		<td  align="center" bgcolor="#FFFFFF"><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
		<td  align="center" bgcolor="#FFFFFF"><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
		<td  align="center" bgcolor="#FFFFFF"><%=db2html(arrList(1,intg))%></td>	
		<td  align="center" bgcolor="#FFFFFF"><%=arrList(2,intg)%></td>	
		<% If etype<>"MD" Then %>
		<td  align="center" bgcolor="#FFFFFF">    
			<a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a>  
		</td>
		<% End If %>
		<td  align="center" bgcolor="#FFFFFF">
			<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrList(0,intg)%>','P')" class="button">
			<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">-->
			<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>','P')"  class="button">
			<% IF arrList(5,intg) = 0 THEN %>
			
			<% 		Response.Write "<a href='" & wwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			<% END IF %>
		</td>					   									
	</tr>
	<%NEXT%>
	</table>
<%END IF%>
</div>
<div id="divIpMG" style="display:none;">
<%IF isArray(arrList) THEN %>
	<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
		<td>�׷��ڵ�</td>					
		<td>�����׷�</td>
		<td>�׷��</td>
		<td>���ļ���</td>
		<% If etype<>"MD" Then %>
		<td>�̹���</td>
		<% End If %>
		<td>����</td>
	</tr>
	<%FOR intg = 0 To UBound(arrList,2)%>				   						
	<tr>
		<td  align="center" bgcolor="#FFFFFF"><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
		<td  align="center" bgcolor="#FFFFFF"><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
		<td  align="center" bgcolor="#FFFFFF"><%=db2html(arrList(1,intg))%></td>	
		<td  align="center" bgcolor="#FFFFFF"><%=arrList(2,intg)%></td>
		<% If etype<>"MD" Then %>
		<td  align="center" bgcolor="#FFFFFF">    
			<a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a>  
		</td>
		<% End If %>
		<td  align="center" bgcolor="#FFFFFF">
			<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrList(0,intg)%>','M')" class="button">
			<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">-->
			<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>','M')"  class="button">
			<% IF arrList(5,intg) = 0 THEN %>
			
			<% 		Response.Write "<a href='" & mobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			<% END IF %>
		</td>					   									
	</tr>
	<%NEXT%>
	</table>
<%END IF%>
</div>
   <script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
	<script type="text/javascript">
		    alert("<%=strMsg%>"); 
		    <%if smode = "F" then%> 
		          $("#divFrm3", parent.document).html($("#divIpG").html()); 
		          parent.document.all.divForm.style.display = "none"; 
		          $("#divMFrm3", parent.document).html($("#divIpMG").html()); 
		           parent.document.all.divForm_mo.style.display = "none";
		    <%elseif smode = "GS" then%>
				<%if eChannel ="P" then%>
					$("#divFrm3", opener.document).html($("#divIpG").html()); 
				<% else %>
					$("#divMFrm3", opener.document).html($("#divIpMG").html()); 
				<%end if%>
		    <%elseif smode = "C" then%>
		        <%if eChannel ="M" then%>
		        $("#divFrm3", parent.document).html($("#divIpG").html()); 
		          parent.document.all.divCopy.style.display = "none";
		           parent.document.all.divForm.style.display = "none";
		       <% else %>
		        $("#divMFrm3", parent.document).html($("#divIpMG").html()); 
		           parent.document.all.divCopy_mo.style.display = "none";
		            parent.document.all.divForm_mo.style.display = "none";
		       <%end if%>
		    <%end if%>
			<%if smode = "GS" then%>
		     self.close();
			<%else%>
			self.location.href = "about:blank";
			<%end if%>
	</script>