<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   


'####################################################
' Page : /admin/eventmanage/event/eventitem_regist.asp
' Description :  �̺�Ʈ ��� - ��ǰ���
' History : 2007.02.21 ������ ����
'           2008.10.20 ��ǰ�̹��� ũ�� �߰�(������)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"--> 
<%
 Dim eCode, itemidarr, mode, sGroup,sortarr, sizearr, sType,sortarr_mo, sizearr_mo, usingarr, usingarr_mo
 Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
 dim tempidarr,cnt,i,sqlStr,strSqlAdd,addSql
 dim eSort,strG
 dim iCurrpage
 dim sgDelivery : sgDelivery = ""
 Dim itemCnt 
 Dim dispCate
 dim using_mo,blnOnlyMobile
 dim eChannel
 dim evtgroup_code_mo, vChangeContents, vSCMChangeSQL

  
mode = Request("mode")

itemidarr = Request("itemidarr")
 
sGroup = trim(Request("selGroup"))
sType =  Request("sType")
 
eCode =request("eC")
itemid      = request("itemid")
itemname    = request("itemname")
makerid     = request("makerid")
sellyn      = request("sellyn")
usingyn     = request("usingyn")
danjongyn   = request("danjongyn")
limityn     = request("limityn")
mwdiv       = request("mwdiv")
sailyn      = request("sailyn")
deliverytype= request("deliverytype")
using_mo	= request("using_mo")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

eChannel = requestCheckvar(request("eCh"),1)
dispCate = requestCheckvar(request("disp"),16) 
iCurrpage = request("iC") 
strG =	 Request("selG")
evtgroup_code_mo  =	 Request("selG_mo")
		  
	dbget.beginTrans
 
Select Case mode
	Case "I" '// ��ǰ�߰� 
	 addSql = ""
	
	'-- ����ǰ������ ���� ��� Ÿ������ üũ-------------------		
	sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup	
	rsget.Open sqlStr, dbget
	IF not rsget.EOF THEN
		sgDelivery = rsget("gift_delivery")
	END IF	
	rsget.close	
	
	IF sgDelivery = "Y" THEN '��ü����� �ܿ�
		 strSqlAdd = " and deliverytype not in (2,5,7,9)"
	ELSE
		strSqlAdd = " and deliverytype not in (1,4)"
	END IF	
	'------------------------------------------------------------
	
	  IF sType = "all" THEN '�˻��� ��� ���� insert  ó��
	  	 '// �߰� ����
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemidarr <> "") then
            addSql = addSql & " and i.itemid in (" + itemidarr + ")"
        end if

        if (itemname <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
        end if
        
        if (sellyn <> "") then
            addSql = addSql & " and i.sellyn='" + sellyn + "'"
        end if

        if (usingyn <> "") then
            addSql = addSql & " and i.isusing='" + usingyn + "'"
        end if
        
        if danjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif danjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
        end if
      		
		if limityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif limityn<>"" then
            addSql = addSql + " and i.limityn='" + limityn + "'"
        end if        
        
        if mwdiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif mwdiv<>"" then
            addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
        end if
		
        if cdl<>"" then
            addSql = addSql + " and i.cate_large='" + cdl + "'"
        end if
        
        if cdm<>"" then
            addSql = addSql + " and i.cate_mid='" + cdm + "'"
        end if
        
        if cds<>"" then
            addSql = addSql + " and i.cate_small='" + cds + "'"
        end If
        
		if dispCate<>"" then
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + dispCate + "%' and isDefault='y') "
		end if
        
        if sailyn<>"" then
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if  
        
         if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if
    ELSE
    	addSql = addSql & " and i.itemid in ("&trim(itemidarr)&")"	    
	END IF	

		'����ǰ�� ���� ��� �̺�Ʈ��ϻ�ǰ  ��� Ȯ��
		IF sgDelivery <> "" THEN
				itemCnt = 0
				sqlStr = " select count(i.itemid) from  [db_item].[dbo].tbl_item i where  1=1 "&addSql	&strSqlAdd					
				rsget.Open sqlStr, dbget
				IF not rsget.EOF THEN
					itemCnt = rsget(0)
				END IF	
				rsget.close	
				
				IF itemCnt > 0 THEN
			%>
				<script language="javascript">
				<!--
				alert("����ǰ�� ������ǰ� �������� ���� ��ǰ�� �߰� �Ұ����մϴ�. ���� Ȯ�� �� �ٽ� ������ּ���");
				self.location.href ="about:blank";
				//-->
				</script>
			<%               
					response.End	
				END IF	
		END IF		 
		
			Dim iChkCount
	 		sqlStr = "SELECT  count(itemid) FROM  [db_item].[dbo].tbl_item as i WHERE itemid not in (select itemid from [db_event].[dbo].tbl_eventitem where evt_code="+eCode+" and evtitem_isUsing = 1) "+addSql 
			rsget.Open sqlStr, dbget
			IF not rsget.EOF THEN
				iChkCount = rsget(0)
			END IF	
			rsget.close	 
			IF iChkCount>1000 THEN
					%>
				<script language="javascript">
				<!--
				alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ���");
				self.location.href ="about:blank";
				//-->
				</script>
			<%               
			response.end
			END IF
		
			sqlStr =" insert into [db_event].[dbo].tbl_eventitem" & VbCrlf
			sqlStr = sqlStr & " (evt_code,itemid,evtgroup_code,evtitem_sort,  evtitem_sort_mo, evtitem_imgsize )" & VbCrlf
			sqlStr = sqlStr & " select " & CStr(eCode)& ", i.itemid, '"&sGroup&"',50,50,153 " & VbCrlf
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i " & VbCrlf
			sqlStr = sqlStr & "     left outer join  [db_event].[dbo].tbl_eventitem  as ei on i.itemid = ei.itemid and ei.evt_code =  "&eCode& VbCrlf
			sqlStr = sqlStr & " where  ei.itemid is null "& addSql  
		 	dbget.execute sqlStr
	    
	    
			sqlStr =" if exists( select ei.itemid from db_item.dbo.tbl_item as i inner join  [db_event].[dbo].tbl_eventitem  as ei on i.itemid = ei.itemid and ei.evt_code =  "&eCode&" and ei.evtitem_isusing = 0 "&addSql&" )" 
			sqlStr = sqlStr & " update ei set evtitem_isusing = 1 "& VbCrlf
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i "  & VbCrlf
			sqlStr = sqlStr & " inner join  [db_event].[dbo].tbl_eventitem  as ei "& VbCrlf
			sqlStr = sqlStr & "  on i.itemid = ei.itemid and ei.evt_code =  "&eCode&" and ei.evtitem_isusing = 0"& addSql 
			dbget.execute sqlStr
			
		
		    
		    ''���̾ ����ǰ �ӽ�..
		    if (CStr(eCode)="8361" or CStr(eCode)="8362" or CStr(eCode)="8363") then
		        sqlStr = "exec db_diary_collection.dbo.ten_IMSI_diary_eventPrize"
		        dbget.execute sqlStr
		    end if
		    
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ������ �߰�. �ڵ� = " & sGroup & vbCrLf
			vChangeContents = vChangeContents & "- ��ǰ�ڵ� = " & trim(itemidarr) & vbCrLf
		 
	Case "D" '// ���û�ǰ ����
			sqlStr = "Delete From  [db_event].[dbo].tbl_eventitem "&_
					"	WHERE evt_code = "&eCode&" and itemid in ("&itemidarr&") "				
			dbget.execute sqlStr
			
		 	
		 	''���̾ ����ǰ �ӽ�..
		    if (CStr(eCode)="8361" or CStr(eCode)="8362" or CStr(eCode)="8363") then
		        sqlStr = "exec db_diary_collection.dbo.ten_IMSI_diary_eventPrize"
		        dbget.execute sqlStr
		    end if
		    
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷� ���û�ǰ ����." & vbCrLf
			vChangeContents = vChangeContents & "- ��ǰ�ڵ� = " & trim(itemidarr) & vbCrLf
			
	Case "G" '//�׷��̵�
		
		'-- ����ǰ������ ���� ��� Ÿ������ üũ-------------------		
			sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup			
			rsget.Open sqlStr, dbget
			IF not rsget.EOF THEN
				sgDelivery = rsget("gift_delivery")
			END IF	
			rsget.close	
			
			IF sgDelivery <> "" THEN
				itemCnt = 0
				IF sgDelivery = "Y" THEN '��ü����� �ܿ�
					strSqlAdd = " and deliverytype not in (2,5,7,9)"
				ELSE
					strSqlAdd = " and deliverytype not in (1,4)"
				END IF					
				
				sqlStr = "SELECT count(itemid) FROM [db_item].[dbo].tbl_item WHERE itemid in  ( "&itemidarr&") " & strSqlAdd			
				rsget.Open sqlStr, dbget
				IF not rsget.EOF THEN
					itemCnt = rsget(0)
				END IF	
				rsget.close	
				
				IF itemCnt > 0 THEN
		%>
			<script language="javascript">
			<!--
			alert("����ǰ ������ ���Ÿ�԰� �ٸ� ��ǰ�� �����մϴ�. �̵� �Ұ����մϴ�.");
			history.back(-1);
			//-->
			</script>
		<% 		dbget.close()	:	response.End
				END IF								
			END IF	
		
		'------------------------------------------------------------			
							
			sqlStr = "UPDATE [db_event].[dbo].tbl_eventitem SET "&_
					" evtgroup_code = "&sGroup&_  
					"	WHERE evt_code = "&eCode&" and itemid in ( "&itemidarr&")  "
			dbget.execute sqlStr
			
			vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") �׷�(" & sGroup & ")���� ���û�ǰ �̵�." & vbCrLf
			vChangeContents = vChangeContents & "- ��ǰ�ڵ� = " & trim(itemidarr) & vbCrLf
					
	Case "S" '//��ǰ����/�̹���ũ�� ����
		Dim tmpSort, tmpSize , tmpDisp,disparr
		sortarr = Request("sortarr")
		sizearr = Request("sizearr") 
		disparr = request("disparr")
		
		If sortarr="" and sizearr=""  and disparr=""  THEN
			dbget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbget.close()	:	response.End
		end if

		'���û�ǰ �ľ�
		itemidarr = split(itemidarr,",")
		cnt = ubound(itemidarr)
         if cnt > 0 then 
        	sortarr =  split(sortarr,",") 
        	sizearr =  split(sizearr,",") 
        	disparr =  split(disparr,",")
        end if			
		'// ���ļ��� ����  
			for i=0 to cnt	
			
    		tmpSort = "NULL"	
    		tmpSize = "NULL"	 
    		tmpDisp = "NULL"		
			    if cnt > 0 then  
    				if sortarr(i)<> "" then	
    				 tmpSort = sortarr(i)	
    				end if 
    			 
    			    if ubound(sizearr) > 0 then
        				if sizearr(i)<> "" then	
        				 tmpSize = sizearr(i)	
        				end if  
    			    end if
    			     
        			if disparr(i)<> "" then	
        			 tmpDisp = disparr(i)	
        			end if
			    else 
    				tmpSort = sortarr 	 
    				tmpSize = sizearr 
    				tmpDisp = disparr	  
		        end if 
					 
			 		 
				sqlStr = "UPDATE [db_event].[dbo].tbl_eventitem SET " 
				if eChannel ="P" then
				sqlStr =	sqlStr&" evtitem_sort = "&tmpSort& " ,evtitem_imgsize = "&tmpSize&", evtitem_isdisp = "& tmpDisp
				sqlStr =	sqlStr&	" ,evtitem_sort_mo = "&tmpSort& ", evtitem_isdisp_mo = "& tmpDisp
				else		
				sqlStr =	sqlStr&	" evtitem_sort_mo = "&tmpSort& ", evtitem_isdisp_mo = "& tmpDisp 
				end if		
				sqlStr =	sqlStr&	"	WHERE evt_code = "&eCode&" and itemid =" & itemidarr(i)    
			 		
				dbget.execute sqlStr
			next   
	
		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") ���û�ǰ ��ǰ����/�̹���ũ�� ����." & vbCrLf
		vChangeContents = vChangeContents & "- itemid = " & trim(Request("itemidarr")) & vbCrLf
		vChangeContents = vChangeContents & "- evtitem_sort = " & trim(Request("sortarr")) & vbCrLf
		vChangeContents = vChangeContents & "- evtitem_imgsize = " & trim(Request("sizearr")) & vbCrLf
		vChangeContents = vChangeContents & "- evtitem_isdisp = " & trim(Request("disparr")) & vbCrLf
	
	Case "L"
	dim eitemlisttype
	eitemlisttype =  requestCheckvar(request("eILT"),2)
	    sqlStr = "UPDATE [db_event].dbo.tbl_event_display SET evt_itemlisttype= '"&eitemlisttype&"'"&_
		"	WHERE evt_code = "&eCode 
		dbget.execute sqlStr 

		vChangeContents = vChangeContents & "- �̺�Ʈ(" & eCode & ") ��ǰ ����Ʈ ��Ÿ�� ����." & vbCrLf
		vChangeContents = vChangeContents & "- evt_itemlisttype = " & eitemlisttype & vbCrLf

End Select
	

	IF Err.Number = 0 THEN
	dbget.CommitTrans

	'### ���� �α� ����(event)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & sGroup & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

	if mode= "I" then
%>
	<script langauge="javascript">
	<!--	
		location.href ="about:blank";
		parent.history.go(0);	
	//-->
	</script>
<%
	else
	    dim strURL	
	    if eChannel = "P" then	
	        strURL = "eventitem_regist.asp"
	    else
	        strURL = "eventitem_regist_mo.asp" 
	    end if
	    
		 Call sbAlertMsg ("����Ǿ����ϴ�.",strURL&"?eC="&eCode&"&menupos="&menupos&"&selG="&strG&"&iC="&iCurrpage&"&makerid="&makerid&"&itemid="&itemid&"&itemname="&itemname&"&chkmo="&blnOnlyMobile, "self") 
		
	end if
	dbget.close()	:	response.End
	Else
   	dbget.RollBackTrans	  
%>
	<script language="javascript">
	<!--
	alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
	history.back(-1);
	//-->
	</script>
<%                
	dbget.close()	:	response.End	
End IF	
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->