<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ��ǰ���
' History : 2008.04.04 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim sMode,sCode, dSDate, dEDate
Dim strSql,addSql, addSqlDB
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype,Keyword,CouponYn
Dim egCode,eCode,itemidarr,sType, i
Dim ErrStr : ErrStr = ""
Dim objCmd,iResult
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus 
Dim dispCate
Dim  dissellprice,disbuyprice,arrsaleItemStatus,saleStatus, tmpsaleItemStatus

sMode     = requestCheckVar(Request("mode"),1)	
SELECT Case sMode
	Case "I"	'���λ�ǰ �߰�
		itemidarr = Request("itemidarr")

		sType 		=  Request("sType")
	
		sCode 		= requestCheckVar(Request("sC"),10)
		eCode 		= requestCheckVar(Request("eC"),10)
		egCode 		= Request("egC")	: if egCode = "" then egCode = 0
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
		Keyword		= request("Keyword")
		CouponYn	= request("CouponYn")
		
		cdl = request("cdl")
		cdm = request("cdm")
		cds = request("cds")
        
        dispCate = requestCheckvar(request("disp"),16)
        
		addSql = ""
		addSqlDB = ""

		IF sType = "all" THEN '�˻��� ��� ���� insert  ó��
			
			'// �߰� ���� 
			if (makerid <> "") then addSql = addSql & " and i.makerid='" & makerid & "'"
 
			if (itemid <> "") then
				dim iA ,arrTemp,arrItemid

				itemid = replace(itemid,chr(13),"") '��ǰ�ڵ�˻� ���ͷ�(2013.12.24)
				arrTemp = Split(itemid,chr(10))
			
				iA = 0
				do while iA <= ubound(arrTemp)
			
					if trim(arrTemp(iA))<>"" then
						'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.04;������)
						if Not(isNumeric(trim(arrTemp(iA)))) then
							Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
							dbget.close()	:	response.End
						else
							arrItemid = arrItemid & trim(arrTemp(iA)) & ","
						end if
					end if
					iA = iA + 1
				loop
			
				itemid = left(arrItemid,len(arrItemid)-1)
				addSql = addSql & " and i.itemid in (" + itemid + ")"
				itemidarr = itemid
			end if
 
			if (itemname <> "") then addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
			
			if (Keyword <> "") then
				addSqlDB = addSqlDB + " Join [db_item].[dbo].tbl_item_contents Ct  on i.itemid=Ct.itemid "
            	addSql = addSql & " and Ct.keywords like '%" + Keyword + "%'"
        	end if	
        	
			if (sellyn <> "") then addSql = addSql & " and i.sellyn='" + sellyn + "'"
			if (usingyn <> "") then addSql = addSql & " and i.isusing='" + usingyn + "'"

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

			if cdl<>"" then addSql = addSql + " and i.cate_large='" + cdl + "'"
			if cdm<>"" then addSql = addSql + " and i.cate_mid='" + cdm + "'"
			if cds<>"" then addSql = addSql + " and i.cate_small='" + cds + "'"
			if dispCate<>"" then
		    	addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + dispCate + "%' and isDefault='y') "
		    end if
		    
			if sailyn<>"" then addSql = addSql + " and i.sailyn='" + sailyn + "'"
			if deliverytype <> "" then addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
			if CouponYn<>"" then  addSql = addSql + " and i.itemCouponyn='" + CouponYn + "'"
        
		ELSE
			addSql = addSql & " and i.itemid in ("&trim(itemidarr)&")"
		END IF

		if eCode <> "" then
			addSqlDB = addSqlDB +  " , [db_event].[dbo].[tbl_eventitem] c "
			addSql = addSql + " and i.itemid = c.itemid and c.evt_code = "&eCode&" and c.evtgroup_code ="&egCode
		end if
		 

		'- ���������� ���� ��ǰ�� ���ؼ� (2013.06.21; MD�� ��û�� ���� ��ü �������� ��ǰ�� �߰� ����)
		'- �߰��Ϸ��� ���αⰣ���� ���ο������� ���� ��ǰ�� ���ؼ�
		dim iSaleRate, iSaleMargin, iSaleMarginValue, iSaleType

		'- �߰��Ϸ��� ���������� �Ⱓ Ȯ��
		strSql = " SELECT convert(varchar(29),sale_startdate,121) as sale_startdate, convert(varchar(29),sale_enddate,121) as sale_enddate, sale_rate, sale_margin, sale_marginvalue, sale_status,sale_type from [db_event].[dbo].tbl_sale where sale_code= "&sCode		
		rsget.Open strSql,dbget 
		IF not rsget.EOF THEN
			dSDate = rsget("sale_startdate")
			dEDate = rsget("sale_enddate")	
			iSaleRate = rsget("sale_rate")	
			iSaleMargin = rsget("sale_margin")	
			iSaleMarginValue = rsget("sale_marginvalue")	
			saleStatus	= rsget("sale_status")
			iSaleType		= rsget("sale_type")
		End IF
		rsget.Close

		dim strStatus, arrList,intLoop
		
		IF itemidarr <> "" THEN
			strSql = "SELECT TOP 1000  b.itemid, a.sale_code, a.sale_status "&_
				"   FROM  [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b "&_
				"   WHERE  a.sale_code = b.sale_code "&_
				"           and  ( "&_ 
				"                    ( ( a.sale_type ='"&iSaleType&"' and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"') "&_
				"	                    and a.sale_using =1 and a.sale_status <> 8 and  b.saleitem_status <> 8 "&_
				"                    ) "&_
				"                    or "&_
				"                    (a.sale_code = "&sCode&")"&_
				"                 ) "&_
				"            and b.itemid in ("&itemidarr&")"  
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				arrList = rsget.getRows()
			End IF
			rsget.Close

			If isArray(arrList) THEN
				For intLoop =0 To UBound(arrList,2)
					Select Case arrList(2,intLoop)
						Case 6
							strStatus = "������"
						Case 7
							strStatus = "���ο���"
						Case 0
							strStatus = "��ϴ��"
					End Select

					ErrStr = ErrStr + "�����ڵ� : " + CStr(arrList(1,intLoop)) + " - ��ǰ��ȣ : " + CStr(arrList(0,intLoop)) +" "+ strStatus + " \n"
				Next
			END IF
		END IF	

		Dim iChkCount,sqlStr
	 		sqlStr = "SELECT  count(i.itemid) FROM  [db_item].[dbo].tbl_item as i " &addSqlDB 
	 		sqlStr = sqlStr &" WHERE i.itemid not in (select itemid from [db_event].[dbo].tbl_saleItem where sale_code="+sCode+") "+addSql  
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
		' ', orgsailprice, orgsailsuplycash, orgsailyn) "
		strSql = "INSERT INTO [db_event].[dbo].[tbl_saleItem]([sale_code], [itemid], [saleItem_status], [saleprice],[salesupplycash])"
		'strSql = strSql&" SELECT "&sCode&", i.itemid, 7, convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100))"
		strSql = strSql&" SELECT "&sCode&", i.itemid, 7, round(convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)), -1, 1)"
		Select Case iSaleMargin
			Case 1		'���ϸ���
				strSql = strSql&" ,convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100))- convert(int,(i.orgprice-(i.orgprice*"&iSaleRate&"/100))*(100-convert(float,convert(int,i.orgsuplycash/i.orgprice*10000)/100))/100)"
			Case 2		'��ü�δ�
				strSql = strSql&" ,convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - (i.orgprice- i.orgsuplycash)"
			Case 3		'�ݹݺδ�
				strSql = strSql&" , i.orgsuplycash - Convert(int, (i.orgprice-(i.orgprice-(i.orgprice*"&iSaleRate&"/100)))/2)"
			Case 4		'10x10�δ�
				strSql = strSql&" , i.orgsuplycash "
			Case 5		'��������
				strSql = strSql&" , convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - convert(int, (i.orgprice-(i.orgprice*"&iSaleRate&"/100))*convert(float,"&iSaleMarginValue&")/100) "
		End Select
		
		'strSql = strSql&"	, i.sailprice, i.sailsuplycash, i.sailyn "
		strSql = strSql&"	FROM [db_item].[dbo].tbl_item i "&addSqlDB
		''''strSql = strSql&"   WHERE i.sailyn ='N' and "			'(2013.06.21; MD�� ��û�� ���� ��ü �������� ��ǰ�� �߰� ����)
		strSql = strSql&" Where i.itemid not in "
		strSql = strSql&" (select b.itemid from [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b"
		strSql = strSql&" 	where a.sale_code = b.sale_code "
		strSql = strSql&" 		and (	( a.sale_type ="&iSaleType&" and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"' and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 )"
		strSql = strSql&"	 			 		or"
		strSql = strSql&"		  			(a.sale_code = "&sCode&")) )"&addSql							
		'response.write strSql
		dbget.execute strSql
	
		IF Err.Number <> 0 THEN			
	       Alert_move "������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���","about:blank" 				
			dbget.close()	:	response.End
		END IF	
		
		IF saleStatus = 6 THEN
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime_v2] ("&sCode&",'"&sMode&"')}"			
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
			 iResult = objCmd(0).Value
			 Set objCmd = nothing
			IF iResult <> 1 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	  
			End IF 
		END IF
%>
	<script type="text/javascript">
	<!--
		<%
		if ErrStr<>"" then
			ErrStr = ErrStr + "\n\n ����Ÿ���� ������ �ߺ����� �Ұ����մϴ�. ���λ�ǰ�� ������ ��ǰ�� �߰��˴ϴ�."
			response.write "alert('" + ErrStr + "')"
		end if
		%>		
		location.href ="about:blank";
		parent.history.go(0);
		//parent.location.reload();	
	//-->
	</script>
<% 	   
		dbget.close()	:	response.End			
	Case "U"	'���� ���û�ǰ ����
		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		itemid 		= split(request("itemid"),",")				
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		saleStatus	=requestCheckVar(Request("saleStatus"),4)
		
		dbget.beginTrans
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
								
				if Cint(trim(arrsaleItemStatus(i))) = 6 then '������ �����϶� �� ����� ���°� ���¿������� ����ó��
					arrsaleItemStatus(i) = 7	
				end if
					 
				IF trim(arrsaleItemStatus(i)) = 9 THEN	'������
					strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
							" SET saleitem_status =9, lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i)) &_
							"	and sale_code=" & sCode
				ELSE	
					strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
							" SET saleprice = "&trim(dissellprice(i))&", salesupplycash="&trim(disbuyprice(i))&", saleitem_status ="&arrsaleItemStatus(i)&", lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i)) &_
							"	and sale_code=" & sCode
				END IF
					dbget.execute strSql
					
				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	  
				End IF
				
				IF Cint(saleStatus) = 6 or Cint(saleStatus) = 9 or Cint(saleStatus) = 8 THEN	'����, ����, ���Ό�� �����϶��� �ٷ� ����	
					Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime_v2] ("&sCode&",'"&sMode&"')}"			
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With	
				    iResult = objCmd(0).Value
				    Set objCmd = nothing
				    
				    IF iResult <> 1 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	  
					End IF
				END IF
				
			end if
		next
		
		dbget.CommitTrans	
		response.redirect("saleReg.asp?menupos="&menupos&"&sC="&sCode&"&eC="&eCode&"&iC="&iCurrpage)
	dbget.close()	:	response.End
	
	Case "R" '���û�ǰ �������� 
	itemid 		= split(request("arrItemid"),",")	
	sCode       = split(request("arrsalecode"),",")	
	
	
	Dim sParm,s_itemid, s_makerid, s_cdl,s_cdm, s_cds, s_dispCate
	s_itemid      = requestCheckvar(request("itemid"),255) 
    s_makerid     = requestCheckvar(request("makerid"),32)
     
    s_cdl = requestCheckvar(request("cdl"),10)
    s_cdm = requestCheckvar(request("cdm"),10)
    s_cds = requestCheckvar(request("cds"),10)
    s_dispCate = requestCheckvar(request("disp"),16)
	sParm = "itemid="&s_itemid&"&makerid="&s_makerid&"&cdl=" &s_cdl&"&cdm=" &s_cdm&"&cds=" &s_cds&"&disp="&s_dispCate 
	
	for i=0 to UBound(itemid) 
			if trim(itemid(i))<>"" then  
			     dbget.beginTrans
					
				strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&vbcrlf
				strSql =strSql &" SET saleitem_status = 8, closedate = getdate(), lastupdate = getdate()"&vbcrlf
				strSql =strSql &" WHERE itemid = "&trim(itemid(i)) &vbcrlf
				strSql =strSql &"	and sale_code=" &trim(sCode(i)) &vbcrlf
				dbget.execute strSql
					
				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	  
				End IF
				
				strSql =" UPDATE i "&vbcrlf
	            strSql =strSql &"SET sellcash = i.orgprice, buycash = i.orgsuplycash, sailprice =0, sailsuplycash = 0, sailyn ='N', mileage=convert(int, i.orgprice*0.005) "&vbcrlf
		        strSql =strSql &"    , availPayType = case a.availPayType when 8 then 0 else  i.availPayType  end "&vbcrlf
		        strSql =strSql &"    , limityn = case  a.availPayType when 8  then b.orglimityn else i.limityn end "&vbcrlf
		        strSql =strSql &"    , lastupdate =getdate() "&vbcrlf
	            strSql =strSql &" FROM [db_item].[dbo].[tbl_item] i,[db_event].[dbo].[tbl_sale] a ,[db_event].[dbo].[tbl_saleItem] b "&vbcrlf
	            strSql =strSql &" WHERE i.itemid = b.itemid and i.sailyn ='Y' and a.sale_code = b.sale_code and (b.orgsailyn = 'N' or b.orgsailyn is null) "&vbcrlf
		        strSql =strSql &"   and a.sale_code =   "&trim(sCode(i))&vbcrlf
		        strSql =strSql &"    and    b.saleitem_status= 8    "&vbcrlf
		        strSql =strSql &"   and b.itemid ="&trim(itemid(i)) 
		      
		        dbget.execute strSql
		    
                IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	 
				End IF
				
            	strSql =" UPDATE i "&vbcrlf
            	strSql =strSql &"SET sellcash = b.orgsailprice, buycash = b.orgsailsuplycash, sailprice =b.orgsailprice, sailsuplycash = b.orgsailsuplycash, sailyn =b.orgsailyn, mileage=convert(int, b.orgsailprice*0.005) "&vbcrlf
            	strSql =strSql &"	, availPayType = case a.availPayType when 8 then 0 else  i.availPayType  end "&vbcrlf
            	strSql =strSql &"	, limityn = case  a.availPayType when 8  then b.orglimityn else i.limityn end "&vbcrlf
            	strSql =strSql &"	, lastupdate =getdate() "&vbcrlf
            	strSql =strSql &"FROM [db_item].[dbo].[tbl_item] i,[db_event].[dbo].[tbl_sale] a ,[db_event].[dbo].[tbl_saleItem] b "&vbcrlf
            	strSql =strSql &"WHERE i.itemid = b.itemid and i.sailyn ='Y' and a.sale_code = b.sale_code and b.orgsailyn = 'Y' "&vbcrlf
            	strSql =strSql &"	and a.sale_code = "&trim(sCode(i))&vbcrlf
            	strSql =strSql &"	   and b.saleitem_status = 8   " &vbcrlf
            	strSql =strSql &"   and b.itemid ="&trim(itemid(i))
            	dbget.execute strSql
            	
                IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		       		dbget.close()	:	response.End	  
				End IF 
  
				dbget.CommitTrans	
				
			end if 
		next
		 
		  response.redirect("saleItemList.asp?"&sParm&"&menupos="&menupos)
		dbget.close()		
	 	response.End
	Case "D"	'���λ�ǰ ����

		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		itemid 		= split(request("itemid"),",")				
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		saleStatus	=requestCheckVar(Request("saleStatus"),4)

		dbget.beginTrans

		for i=0 to UBound(itemid)-1
			if trim(arrsaleItemStatus(i))="9" and trim(itemid(i))<>"" then	'������
				strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
						" SET saleitem_status=9, lastupdate=getdate()"&_
						" WHERE itemid = "&trim(itemid(i)) &_
						" and sale_code=" & sCode
				dbget.execute strSql
			end if
			IF Err.Number <> 0 THEN
				dbget.RollBackTrans	
				Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
				dbget.close()	:	response.End	  
			End IF
		next

		strSql ="EXEC [db_item].[dbo].[sp_Ten_item_SetSaleDeleteItemOrgPrice_RealTime] " & sCode
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans	
			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
			dbget.close()	:	response.End	  
		End IF

		dbget.CommitTrans
		response.redirect("saleReg.asp?menupos="&menupos&"&sC="&sCode&"&eC="&eCode)
		dbget.close()	:	response.End
	Case "P"	'��ǰ���̺� ����
		sCode = requestCheckVar(Request("sC"),10)	
		iCurrpage 	= request("iC")
		IF sCode = "" THEN 
			Alert_return("�Ķ���Ͱ��� ������ �ֽ��ϴ�.")    	
		     dbget.close()	:	response.End	
		END IF       		
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime_v2] ("&sCode&",'"&sMode&"')}"			
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With	
	    iResult = objCmd(0).Value
	    Set objCmd = nothing
	  
	'�˻��� üũ--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'�˻���	
	 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	 sEdate     	= requestCheckVar(Request("iED"),10)		'������	
	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'���� ������ ��ȣ
	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'�˻� ����
 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
 	'--------------------------------------------------------------
 
		IF iResult <> 1 THEN
			Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")    	
		    dbget.close()	:	response.End	
		END IF
			response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
	dbget.close()	:	response.End
	CASE Else
	Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���2")    	
       dbget.close()	:	response.End
END SELECT
	

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
