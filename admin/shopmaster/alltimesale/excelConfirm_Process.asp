<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim itemid, salePrice, saleBuyPrice
Dim strSql, i, ChkCnt, adminid, tmpArrItemid, spiltItemid

adminid  = session("ssBctId")
tmpArrItemid = request("tmpArrItemid")
If Right(tmpArrItemid,1) = "," Then
	tmpArrItemid = Left(tmpArrItemid, Len(tmpArrItemid) - 1)
End If
spiltItemid = Split(tmpArrItemid, ",")
ChkCnt = Ubound(spiltItemid)

strSql = "create table #tmpATSale(itemid int, sellprice money, buyprice money)"
For i = 0 To ChkCnt
	itemid = getNumeric(spiltItemid(i))
	salePrice = getNumeric(request("salePrice_" + CStr(itemid)))
	saleBuyPrice = getNumeric(request("saleBuyPrice_" + CStr(itemid)))
	strSql = strSql & " INSERT INTO #tmpATSale (itemid, sellprice, buyprice ) VALUES ('"& itemid &"', '"& salePrice &"', '"& saleBuyPrice &"'); " & VBCRLF
Next
dbget.execute strSql

''이벤트 할인 있을경우 종료처리
strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
strSql = strSql & " (  select si.itemid, 2, s.sale_code, si.saleprice, si.salesupplycash,4,'할인종료','"&adminid&"'"
strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "
strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "
strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
strSql = strSql & "              	and s.sale_using =1   "
strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  <> 8  )"
dbget.execute strSql

strSql = " update si  SET saleitem_status = 9 ,closedate=getdate(), lastupdate =getdate()"
strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "
strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "
strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
strSql = strSql & "              	and s.sale_using =1   "
strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  <> 8  "
dbget.execute strSql

'just1day 상품 확인
strSql ="select si.itemid  "
strSql = strSql & " into #tmpJ1day "
strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "
strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "
strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
strSql = strSql & "              	and s.sale_using =1   "
strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  "
dbget.execute strSql

strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
strSql = strSql & " (  select si.itemid, 2, s.sale_code, si.saleprice, si.salesupplycash,5,'저스트원데이 할인중-상시할인등록대기처리','"&adminid&"'"
strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "
strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "
strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
strSql = strSql & "              	and s.sale_using =1   "
strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  )"
dbget.execute strSql

strSql = " update si  SET orgsailprice = t.sellprice,orgsailsuplycash = t.buyprice , orgsailyn='Y',lastupdate =getdate()"
strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "
strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "
strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
strSql = strSql & "              	and s.sale_using =1   "
strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  "
dbget.execute strSql

' 상시할인 처리
strSql = "update  i "
strSql = strSql & " set sellcash = t.sellprice, buycash = t.buyprice, sailprice =t.sellprice  , sailsuplycash =t.buyprice , sailyn ='Y'"
strSql = strSql & " , mileage=case when (1-(convert(float,t.sellprice)/ i.orgprice)) >= 0.4 then 0 else convert(int, t.sellprice*0.005) end, lastupdate =getdate()"
strSql = strSql & " from db_item.dbo.tbl_item as i "
strSql = strSql & " inner join #tmpATSale as t on i.itemid  = t.itemid "
strSql = strSql & " left outer join #tmpJ1day as j on t.itemid = j.itemid "
strSql = strSql & " where j.itemid is null "
dbget.execute strSql

strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
strSql = strSql &" select t.itemid ,1,0, t.sellprice,t.buyprice,1,'상시할인등록','"&adminid&"' "
strSql = strSql & " from #tmpATSale as t  "
strSql = strSql & " left outer join #tmpJ1day as j on t.itemid = j.itemid "
strSql = strSql & " where j.itemid is null "
dbget.execute strSql

'임시테이블 삭제
strSql = "drop table #tmpJ1day"
dbget.execute strSql

strSql = "drop table #tmpATSale "
dbget.execute strSql
response.write "<script>top.opener.opener.location.reload();top.opener.window.close();top.window.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->