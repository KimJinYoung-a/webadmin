<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ��ǰ���
' History : 2008.04.04 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
 Dim gCode, itemidarr, mode, sGroup,sortarr, sType, sgDelivery
 Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,deliverytype
 dim tempidarr,cnt,i,sqlStr,addSql,strSqlAdd
  dim iCurrpage
 
mode = Request("mode")

itemidarr = Request("itemidarr")
 

sType =  Request("sType")
 
gCode =request("gC")
itemid      = request("itemid")
itemname    = request("itemname")
makerid     = request("makerid")
sellyn      = request("sellyn")
usingyn     = request("usingyn")
danjongyn   = request("danjongyn")
limityn     = request("limityn")
deliverytype= request("deliverytype")
sailyn      = request("sailyn")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

iCurrpage = request("iC") 

Select Case mode
	Case "I" '// ��ǰ�߰� 
	 addSql = ""
	 
	 '-- ����ǰ������ ���� ��� Ÿ������ üũ-------------------
	sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_code = "&gCode
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
	
		Dim itemCnt : itemCnt = 0
		sqlStr = " select count(i.itemid) from  [db_item].[dbo].tbl_item i where  1=1 "&addSql&strSqlAdd		
		rsget.Open sqlStr, dbget
		IF not rsget.EOF THEN
			itemCnt = rsget(0)
		END IF	
		rsget.close	
		
		IF itemCnt > 0 THEN
			Call Alert_move("����ǰ�� ������ǰ� �������� ���� ��ǰ�� �߰� �Ұ����մϴ�. ���� Ȯ�� �� �ٽ� ������ּ���","about:blank")				
			dbget.close()	:	response.End	
		END IF	
			
			sqlStr = " insert into [db_event].[dbo].tbl_giftitem" + VbCrlf
			sqlStr = sqlStr + " (gift_code,itemid)" + VbCrlf
			sqlStr = sqlStr + " select " + CStr(gCode) + ", i.itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " itemid not in "
			sqlStr = sqlStr + " (select itemid from [db_event].[dbo].tbl_giftitem"
			sqlStr = sqlStr + " where gift_code=" + gCode
			sqlStr = sqlStr + " )"	 + addSql							
			dbget.execute sqlStr
			 
	Case "D" '// ���û�ǰ ����
			sqlStr = "Delete From  [db_event].[dbo].tbl_giftitem "&_
					"	WHERE gift_code = "&gCode&" and itemid in ("&itemidarr&") "				
			dbget.execute sqlStr 		
End Select
	

	IF Err.Number = 0 THEN
		if mode= "I" then
	%>
	<script langauge="javascript">
	<!--	
		location.href ="about:blank";
		parent.history.go(0);	
	//-->
	</script>
<% 	   
		dbget.close()	:	response.End		
		response.redirect("giftItemReg.asp?gC="&gCode&"&menupos="&menupos&"&iC="&iCurrpage)
		end if
		dbget.close()	:	response.End
	Else
   
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