<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������� ���� ������ó��
' History : 2018.02.12 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim sMode
dim arritemid, i, itemid
dim arrsell, arrbuy, sellcash, buycash
dim strSql
dim isCheck
dim arrList, intLoop
dim adminid
dim menupos,isJust
sMode     = requestCheckVar(Request("hidM"),1)	
menupos     = requestCheckVar(Request("menupos"),32)	
adminid  = session("ssBctId")
SELECT CASE sMode
	CASE "S" '������ε��
	isCheck = False
	  	arritemid 		= split(request("itemid"),",")		
	  	arrsell 		= split(request("iDSPrice"),",")		
	  	arrbuy 		= split(request("iDBPrice"),",")		
	  	For i = 0  To UBOund(arritemid)-1
	  	  itemid = requestCheckVar(trim(arritemid(i)),10)
	  	  sellcash= requestCheckVar(trim(arrsell(i)),30)
	  	  buycash= requestCheckVar(trim(arrbuy(i)),30)
	  	  if  (isNumeric(itemid)) and (isNumeric(sellcash)) and (isNumeric(buycash)) then'��ǰ�ڵ�, �ǸŰ�, ���ް� ���ڿ��� üũ 
	  	  	isCheck = True
	  	  else
	  	  	isCheck = False		
	  	  end if
	  	  
	  	  if sellcash > 0 and buycash > 0 then
	  	  	isCheck = True
	  	  else
	  	  	isCheck = False	
	  	  end if	
	  	
	  	  	if  isCheck then 
	  	  		'�̺�Ʈ ���� ������� ����ó��
	  	  		'just1day�� ���� ���� ?!
	  	  		 strSql = " select s.sale_code, si.itemid, si.saleprice, si.salesupplycash,s.availPayType  FROM db_event.dbo.tbl_sale as s   "
	           strSql = strSql & "              inner join 	db_event.dbo.tbl_saleitem as si on s.sale_code = si.sale_code   "	           
	           strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
	           strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
	           strSql = strSql & "              	and s.sale_using =1   "	           
	           strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121)  "
	           strSql = strSql & "								and si.itemid = "&itemid &""
	           rsget.Open strSql,dbget
						IF not rsget.EOF THEN
							 arrList = rsget.getRows()
						END IF
						rsget.close
						
						if isArray(arrList) THEN
							for intLoop = 0 To ubound(arrList,2)
							if arrList(4,intLoop) = 8 THEN 'just1day ��ǰ �� ���� ���ó��
								 strSql = " UPDATE  db_event.dbo.tbl_saleitem "
	           		strSql = strSql & " SET orgsailprice ="&sellcash&",orgsailsuplycash ="&buycash&", orgsailyn ='Y', lastupdate =getdate()"
	           		strSql = strSql & "	where itemid = "&arrList(1,intLoop)&" and sale_code = "&arrList(0,intLoop)
	           		dbget.execute strSql
	           			
	           		strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	           		strSql = strSql & " values("&arrList(1,intLoop)&",2,"&arrList(0,intLoop)&",'"&arrList(2,intLoop)&"','"&arrList(3,intLoop)&"',5,'����Ʈ������ ������-������ε�ϴ��ó��','"&adminid&"')"
	           		dbget.execute strSql
	           		
	           		isJust = arrList(4,intLoop)
							else	
							  strSql = " UPDATE  db_event.dbo.tbl_saleitem "
	           		strSql = strSql & " SET saleitem_status = 9 ,closedate=getdate(), lastupdate =getdate()"
	           		strSql = strSql & "	where itemid = "&arrList(1,intLoop)&" and sale_code = "&arrList(0,intLoop)
	           		dbget.execute strSql
	           			
	           		strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	           		strSql = strSql & " values("&arrList(1,intLoop)&",2,"&arrList(0,intLoop)&",'"&arrList(2,intLoop)&"','"&arrList(3,intLoop)&"',4,'��������','"&adminid&"')"
	           		dbget.execute strSql
	           	end if	
							next 
							
					  end if
					  
					   
					 		if isJust <> "8" THEN  'just1day �ƴҶ��� ��ǰ �ٷ� ����ó��	  		 
						 strSql = "update  db_item.dbo.tbl_item  "
						 strSql = strSql & " set sellcash = '"&sellcash&"', buycash = '"&buycash&"', sailprice = '"&sellcash&"', sailsuplycash = '"&buycash&"', sailyn ='Y'"
						 strSql = strSql & " , mileage=case when (1-(convert(float,"&sellcash&")/ orgprice)) >= 0.4 then 0 else convert(int, "&sellcash&"*0.005) end, lastupdate =getdate()"
						 strSql = strSql & " where itemid = "&itemid
						 dbget.execute strSql
						 
						 strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	           		strSql = strSql &" values("&itemid&",1,0,'"&sellcash&"','"&buycash&"',1,'������ε��','"&adminid&"')"
	           		 dbget.execute strSql
	           		end if
					end if
	  	Next
	  	
	   Alert_move "ó���Ǿ����ϴ�.","/admin/shopmaster/alltimesale/?menupos="&menupos 
			dbget.close()	:	response.End
	CASE "O" '���� ����
	dim chkJ1day
	chkJ1day = False
	 	arritemid 		= split(request("itemid"),",")
	 	 	For i = 0  To UBOund(arritemid)-1
	  	  itemid = requestCheckVar(trim(arritemid(i)),10)
	  	  
	  	 
	  	     '�̺�Ʈ ��������
	  	     strSql = " select s.sale_code, si.itemid, si.saleprice, si.salesupplycash, s.availPayType "
	  	     strSql = strSql & " from db_event.dbo.tbl_sale as s "
	  	     strSql = strSql & "   inner join 	db_event.dbo.tbl_saleitem as si on s.sale_code = si.sale_code   "	   
	  	     strSql = strSql & " where si.itemid ="&itemid &" and si.saleitem_status <> 9 "
	  	       rsget.Open strSql,dbget
						IF not rsget.EOF THEN
							 arrList = rsget.getRows()
						END IF
						rsget.close
						
						if isArray(arrList) THEN
								for intLoop =0 To ubound(arrList,2)
									if arrList(4,intLoop) = 8 THEN  'just1day ����� ������� �ƴ� ������ ����ó��
										chkJ1day = True 
	           					strSql = " update si " 
	           					strSql = strSql & " SET orgsailprice = i.orgprice ,orgsailsuplycash = i.orgsuplycash, orgsailyn ='N', lastupdate =getdate()"
	           					strSql = strSql & " FROM  db_event.dbo.tbl_saleitem  as si "
	           					strSql = strSql & "	inner join db_item.dbo.tbl_item as i on si.itemid = i.itemid  "
	           					strSql = strSql & " and si.sale_code ="&arrList(0,intLoop)&" and si.itemid ="&arrList(1,intLoop)&" and si.saleitem_status <> 9 "
	           						dbget.execute strSql
	           						
	           						strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) " 
	           					strSql = strSql & " values("&arrList(1,intLoop)&",2,"&arrList(0,intLoop)&",'"&arrList(2,intLoop)&"','"&arrList(3,intLoop)&"',5,'����Ʈ������ ������-������� ����ó��','"&adminid&"')"
	           					dbget.execute strSql
									else
										
										'�̺�Ʈ��������
										 strSql = "update db_event.dbo.tbl_saleitem  "
										strSql = strSql & " SET saleitem_status = 9 ,closedate=getdate(), lastupdate =getdate()"
	           				strSql = strSql & "	where itemid = "&arrList(1,intLoop)&" and sale_code = "&arrList(0,intLoop)
	           				dbget.execute strSql
	           				
	           				strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	           				strSql = strSql & " values("&arrList(1,intLoop)&",2,"&arrList(0,intLoop)&",'"&arrList(2,intLoop)&"','"&arrList(3,intLoop)&"',4,'��������','"&adminid&"')"
	           				dbget.execute strSql
	           	
	           		end if
							next
						end if 	
										'��ǰ�����������
											strSql = "update  i  "
											strSql = strSql & " set sellcash = i.orgprice, buycash = i.orgsuplycash, sailprice = 0, sailsuplycash = 0, sailyn ='N', mileage=convert(int, i.orgprice*0.005), lastupdate =getdate()"
											strSql = strSql & " from db_item.dbo.tbl_item as i "
											strSql = strSql & " where itemid = "&itemid
											dbget.execute strSql
											 
							 				strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
						          strSql =  strSql &  " ( select itemid ,1,0,sellcash, buycash,2,'�����������','"&adminid&"'"
						           strSql =  strSql & " from db_item.dbo.tbl_item where itemid = "&itemid&"   )"
						          dbget.execute strSql
	           					
								
	
			Next
			
			Alert_move "ó���Ǿ����ϴ�.","/admin/shopmaster/alltimesale/?menupos="&menupos 
			dbget.close()	:	response.End
CASE ELSE
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���2")    	
       dbget.close()	:	response.End
END SELECT
%>