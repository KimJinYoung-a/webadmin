<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ ��� - ��ǰ���
' History : 2010.09.29 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<%
Dim eCode, itemidarr, mode, sGroup,sortarr, sizearr, sType
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
dim tempidarr,cnt,i,sqlStr,strSqlAdd,addSql ,eSort,strG ,iCurrpage ,itemCnt
dim sgDelivery : sgDelivery = ""
	mode = RequestCheckvar(Request("mode"),2)
	itemidarr = Request("itemidarr")
	sGroup = RequestCheckvar(Request("selGroup"),10)
	sType =  RequestCheckvar(Request("sType"),10)
	eCode = RequestCheckvar(request("eC"),10)
	itemid      = RequestCheckvar(request("itemid"),10)
	itemname    = RequestCheckvar(request("itemname"),64)
	makerid     = RequestCheckvar(request("makerid"),32)
	sellyn      = RequestCheckvar(request("sellyn"),1)
	usingyn     = RequestCheckvar(request("usingyn"),1)
	danjongyn   = RequestCheckvar(request("danjongyn"),1)
	limityn     = RequestCheckvar(request("limityn"),1)
	mwdiv       = RequestCheckvar(request("mwdiv"),1)
	sailyn      = RequestCheckvar(request("sailyn"),1)
	deliverytype= RequestCheckvar(request("deliverytype"),2)
	cdl = RequestCheckvar(request("cdl"),10)
	cdm = RequestCheckvar(request("cdm"),10)
	cds = RequestCheckvar(request("cds"),10)
	iCurrpage = RequestCheckvar(request("iC"),10)
	strG =	 RequestCheckvar(Request("selG"),10)
  	if itemidarr <> "" then
		if checkNotValidHTML(itemidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
dbacademyget.beginTrans
 
Select Case mode
	
	Case "I" '// ��ǰ�߰� 
	 addSql = ""
	
	'-- ����ǰ������ ���� ��� Ÿ������ üũ-------------------		
	sqlStr = "SELECT gift_delivery FROM [db_academy].[dbo].tbl_gift WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup	
	
	'response.write sqlStr &"<br>"
	rsacademyget.Open sqlStr, dbacademyget
	
	IF not rsacademyget.EOF THEN
		sgDelivery = rsacademyget("gift_delivery")
	END IF	
	rsacademyget.close	
	
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

        if (itemid <> "") then
            addSql = addSql & " and i.itemid in (" + itemid + ")"
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
        end if
        
        if sailyn<>"" then
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if  
        
         if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if
    ELSE
    	addSql = addSql & " and itemid in ("&trim(itemidarr)&")"	    
	END IF	
		'����ǰ�� ���� ��� �̺�Ʈ��ϻ�ǰ  ��� Ȯ��
		IF sgDelivery <> "" THEN
				itemCnt = 0
				sqlStr = " select count(i.itemid) from [db_academy].dbo.tbl_diy_item i where  1=1 "&addSql	&strSqlAdd					
				
				'response.write sqlStr &"<br>"
				rsacademyget.Open sqlStr, dbacademyget
				
				IF not rsacademyget.EOF THEN
					itemCnt = rsacademyget(0)
				END IF	
				rsacademyget.close	
				
				IF itemCnt > 0 THEN
			%>
				<script language="javascript">

				alert("����ǰ�� ������ǰ� �������� ���� ��ǰ�� �߰� �Ұ����մϴ�. ���� Ȯ�� �� �ٽ� ������ּ���");
				self.location.href ="about:blank";

				</script>
			<%               
					response.End	
				END IF	
		END IF		 
		
			sqlStr = " insert into [db_academy].[dbo].tbl_eventitem" + VbCrlf
			sqlStr = sqlStr + " (evt_code,itemid,evtgroup_code,evtitem_sort)" + VbCrlf
			sqlStr = sqlStr + " select " + CStr(eCode) + ", i.itemid, '"&sGroup&"',50"
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " itemid not in ("
			sqlStr = sqlStr + " 	select itemid from [db_academy].[dbo].tbl_eventitem"
			sqlStr = sqlStr + " 	where evt_code=" + eCode
			sqlStr = sqlStr + " )"	 + addSql				
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
	    		    
	Case "D" '// ���û�ǰ ����
			sqlStr = "Delete From [db_academy].[dbo].tbl_eventitem "&_
					" WHERE evt_code = "&eCode&" and itemid in ("&itemidarr&") "				
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
	
	Case "G" '//�׷��̵�		
		'-- ����ǰ������ ���� ��� Ÿ������ üũ-------------------		
			sqlStr = "SELECT gift_delivery FROM [db_academy].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup			
			
			'response.write sqlStr &"<br>"
			rsacademyget.Open sqlStr, dbacademyget
			
			IF not rsacademyget.EOF THEN
				sgDelivery = rsacademyget("gift_delivery")
			END IF	
			
			rsacademyget.close	
			
			IF sgDelivery <> "" THEN
				itemCnt = 0
				IF sgDelivery = "Y" THEN '��ü����� �ܿ�
					strSqlAdd = " and deliverytype not in (2,5,7,9)"
				ELSE
					strSqlAdd = " and deliverytype not in (1,4)"
				END IF					
				
				sqlStr = "SELECT count(itemid) FROM [db_academy].dbo.tbl_diy_item WHERE itemid in  ( "&itemidarr&") " & strSqlAdd			
				
				'response.write sqlStr &"<br>"
				rsacademyget.Open sqlStr, dbacademyget
				
				IF not rsacademyget.EOF THEN
					itemCnt = rsacademyget(0)
				END IF	
				
				rsacademyget.close	
				
				IF itemCnt > 0 THEN
		%>
			<script language="javascript">

			alert("����ǰ ������ ���Ÿ�԰� �ٸ� ��ǰ�� �����մϴ�. �̵� �Ұ����մϴ�.");
			history.back(-1);

			</script>
		<% 		dbacademyget.close()	:	response.End
				END IF								
			END IF	
		
		'------------------------------------------------------------			
							
			sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET"&_
					" evtgroup_code = "&sGroup& _
					" WHERE evt_code = "&eCode&" and itemid in ( "&itemidarr&") "
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
					
	Case "S" '//��ǰ����/�̹���ũ�� ����
		Dim tmpSort, tmpSize
		sortarr = Request("sortarr")
		sizearr = Request("sizearr")

		If sortarr="" and sizearr="" THEN
			dbacademyget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbacademyget.close()	:	response.End
		end if

		'���û�ǰ �ľ�
		itemidarr = split(itemidarr,",")
		cnt = ubound(itemidarr)

		'// ���ļ��� ����
		If sortarr<>"" THEN
			sortarr =  split(sortarr,",")
			
			for i=0 to cnt	
				IF sortarr(i) = "" THEN
					 tmpSort = "NULL"				
				ELSE	
					 tmpSort = sortarr(i)	
				END IF	 
				sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET "&_
						" evtitem_sort = "&tmpSort& _
						" WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				
				'response.write sqlStr &"<br>"
				dbacademyget.execute sqlStr
			next
		END IF

		'// �̹��� ũ�� ����
		If sizearr<>"" THEN			
			sizearr =  split(sizearr,",")

			for i=0 to cnt	
				IF sizearr(i) = "" THEN
					 tmpSize = "NULL"				
				ELSE	
					 tmpSize = sizearr(i)	
				END IF	 
				sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET "&_
						" evtitem_imgsize = "&tmpSize& _
						" WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				
				'response.write sqlStr &"<br>"
				dbacademyget.execute sqlStr
			next
		End If

End Select
	
IF Err.Number = 0 THEN
	dbacademyget.CommitTrans

	if mode= "I" then
%>
		<script langauge="javascript">
	
			location.href ="about:blank";
			parent.history.go(0);	

		</script>
<%
		else		
			response.redirect("eventitem_regist.asp?eC="&eCode&"&menupos="&menupos&"&selG="&strG&"&iC="&iCurrpage)
		end if
	dbacademyget.close()	:	response.End
Else
   	dbacademyget.RollBackTrans	  
%>
	<script language="javascript">
	
	alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
	history.back(-1);
	
	</script>
<%                
	dbacademyget.close()	:	response.End	
End IF	
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->