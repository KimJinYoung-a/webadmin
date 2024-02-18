<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

'' 한정수량 조정 
function LimitedItemCheck(byval orderserial,byval iid,byval iopt,byval flag)
        dim sqlStr
        dim buf, tmp, i
        dim limityn, limitsold
        dim itemid

		if (flag=false) then exit function

		'' 한정판매 Item..
		sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
		sqlStr = sqlStr + " set limitsold=(case when limitno<limitsold + T.itemno then limitno else limitsold + T.itemno end) " + vbCrlf
		sqlStr = sqlStr + " from " + vbCrlf
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select d.itemid, sum(d.itemno) as itemno" + vbCrlf
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d " + vbCrlf
		sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
		sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
		sqlStr = sqlStr + " 	group by d.itemid" + vbCrlf
		sqlStr = sqlStr + " ) as T" + vbCrlf
		sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"

		dbget.Execute(sqlStr)


		''옵션있는상품 - 한정수량이 마이너스가 되지 않게..
		sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
		sqlStr = sqlStr + " set optlimitsold=(case when optlimitno<optlimitsold+T.itemno then optlimitno else optlimitsold+T.itemno end) " + vbCrlf
		sqlStr = sqlStr + " from " + vbCrlf
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select d.itemid, d.itemoption, d.itemno" + vbCrlf
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d " + vbCrlf
		sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
		sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
		sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
		sqlStr = sqlStr + " ) as T" + vbCrlf
		sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

		dbget.Execute(sqlStr)

end function



Const DELIMROW = "Y|R|T"
Const DELIMCOL = "Y|C|T"

dim RowData, rowcount, ColData, i, ParsingRowdata, rcnt
dim ORGData, extorderserial
ORGData         = request("ORGData")
extorderserial  = request("extorderserial")


if (Len(extorderserial)<10) then
    response.write "<script>alert('제휴사 원 주문 번호가 올바르지 않습니다. ');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End    
end if

RowData = split(ORGData,DELIMROW)

if IsArray(rowdata) then
    rowcount = UBound(rowdata) + 1
else
    rowcount = 0
end if


rcnt = 0
for i=0 to rowcount-1 
    ColData = split(Trim(RowData(i)),DELIMCOL)
    
    if ((ColData(4))=extorderserial) then
        rcnt = rcnt + 1
    end if
next

if (rcnt>0) then
    redim ParsingRowdata(rcnt)
    rcnt = 0
    for i=0 to rowcount-1 
        ColData = split(Trim(RowData(i)),DELIMCOL)
        if ((ColData(4))=extorderserial) then
            rcnt = rcnt + 1
            ParsingRowdata(rcnt-1) =  RowData(i)
        end if
    next
end if
        
rowcount = rcnt

'response.write extorderserial & "<br>"
'for i=0 to rowcount-1 
'    response.write ParsingRowdata(i) & ".<br>"
'next 

if (rowcount<1) then
    response.write "<script>alert('데이터가 올바르지 않습니다. ');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End   
end if

dim sqlStr, tenorderserial

''기존 입력 내역 체크
sqlStr = "select orderserial from [db_order].[dbo].tbl_order_master"
sqlStr = sqlStr + " where sitename='yeoin'"
sqlStr = sqlStr + " and cancelyn='N'"
sqlStr = sqlStr + " and authcode='" + extorderserial + "'"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    tenorderserial = rsget("orderserial")
end if
rsget.Close

if (tenorderserial<>"") then
    response.write "<script>alert('이미 저장된 주문입니다. - 주문번호 : " + tenorderserial + " ');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End    
end if


dim tmpItemList, distinctrealcnt

''상품번호 체크

for i=0 to rowcount-1 
	ColData = split(ParsingRowdata(i),DELIMCOL)
	tmpItemList = tmpItemList + ColData(6) + "," 
next

if (Right(tmpItemList,1)=",") then tmpItemList=Left(tmpItemList,Len(tmpItemList)-1)
	
'sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item"
'sqlStr = sqlStr + " where itemid in (" + tmpItemList + ")"
'
'rsget.Open sqlStr, dbget, 1
'    distinctrealcnt = rsget("cnt")
'rsget.Close
'
'if (rowcount<>distinctrealcnt) then
'	'response.write "<script>alert('알수 없는 상품 코드가 있습니다." + CStr(rowcount) + "<>" + CStr(distinctrealcnt) + "')</script>"
'	'response.write "<script>history.back();</script>"
'	'dbget.close()	:	response.End
'end if
	

''order_master 입력
dim iid, orderserial
dim buyname, buyphone, buyhp, buyemail, reqname, reqzipcode
dim reqaddress, reqphone, reqhp, comment, subtotalprice, reqzipaddr
dim buf, deliverpay

ColData = split(ParsingRowdata(0),DELIMCOL)
reqname     = ColData(13)
reqphone    = ColData(14)
reqhp       = ColData(15)

reqzipcode  = ColData(16) 
buf         = ColData(17) '':주소 ALL
reqzipaddr  = SplitValue(buf," ",0) & " " & SplitValue(buf," ",1)
reqaddress  = Trim(Mid(buf,Len(reqzipaddr)+1,250))


comment     = ColData(18)
deliverpay  = ColData(19)

buyname     = ColData(23)
buyphone    = ColData(24)
buyhp       = ColData(25)
subtotalprice = 0




sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
rsget.Open sqlStr,dbget,1,3
rsget.AddNew
rsget("orderserial") = ""
rsget("jumundiv") = "5"
rsget("userid") = ""
rsget("accountname") = ""
rsget("accountdiv") = "50"
rsget("authcode") = extorderserial
rsget.update
    iid = rsget("idx")
rsget.close
	
orderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
orderserial = orderserial & Format00(5,Right(CStr(iid),5))


sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
sqlStr = sqlStr + " set orderserial='" + CStr(orderserial) + "'," & vbCrlf
sqlStr = sqlStr + " accountname='" + html2db(buyname) + "'," & vbCrlf
sqlStr = sqlStr + " totalcost=subtotalprice," & vbCrlf
sqlStr = sqlStr + " totalsum=subtotalprice," & vbCrlf
sqlStr = sqlStr + " ipkumdiv='4'," & vbCrlf
sqlStr = sqlStr + " ipkumdate=getdate()," & vbCrlf
sqlStr = sqlStr + " regdate=getdate()," & vbCrlf
sqlStr = sqlStr + " beadaldiv='1'," & vbCrlf
sqlStr = sqlStr + " buyname='" + html2db(buyname) + "'," & vbCrlf
sqlStr = sqlStr + " buyphone='" + replace(buyphone,"'","") + "'," & vbCrlf
sqlStr = sqlStr + " buyhp='" + replace(buyhp,"'","") + "'," & vbCrlf
sqlStr = sqlStr + " buyemail='" + html2db(buyemail) + "'," & vbCrlf
sqlStr = sqlStr + " reqname='" + html2db(reqname) + "'," & vbCrlf
sqlStr = sqlStr + " reqzipcode='" + replace(reqzipcode,"'","") + "'," & vbCrlf
sqlStr = sqlStr + " reqaddress='" + html2db(reqaddress) + "'," & vbCrlf
sqlStr = sqlStr + " reqphone='" + replace(reqphone,"'","") + "'," & vbCrlf
sqlStr = sqlStr + " reqhp='" + replace(reqhp,"'","") + "'," & vbCrlf
sqlStr = sqlStr + " comment='" + html2db(comment) + "'," & vbCrlf
sqlStr = sqlStr + " sitename='yeoin'," & vbCrlf
sqlStr = sqlStr + " discountrate=1.00," & vbCrlf
sqlStr = sqlStr + " subtotalprice=" + Cstr(subtotalprice) + "," & vbCrlf
sqlStr = sqlStr + " reqzipaddr='" + html2db(reqzipaddr) + "'" & vbCrlf
sqlStr = sqlStr + " where idx=" + CStr(iid)


dbget.Execute sqlStr

''order_detail 입력
dim itemid, itemoption, itemno, itemcost

''배송비
itemid = 0
if (deliverpay=0) then
    itemoption = "0501" 
else
    itemoption = "0000"
end if
itemno = 1
itemcost = deliverpay 

sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
sqlStr = sqlStr + "itemoption, itemno, itemcost, itemvat, mileage, costtotal, orderdate)" & vbCrlf
sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
sqlStr = sqlStr + " '" + orderserial + "'," & vbCrlf
sqlStr = sqlStr + " " + CStr(itemid) + "," & vbCrlf
sqlStr = sqlStr + " '" + itemoption + "'," & vbCrlf
sqlStr = sqlStr + " " + CStr(itemno) + "," & vbCrlf
sqlStr = sqlStr + "	" + CStr(itemcost) + "," & vbCrlf
sqlStr = sqlStr + "	0,0,0,getdate()" & vbCrlf
sqlStr = sqlStr + " )"


dbget.Execute sqlStr

''상품 입력

for i=0 to rowcount-1 
	ColData = split(ParsingRowdata(i),DELIMCOL)
	
    itemid      = ColData(6)
    itemoption  = Trim(ColData(9))
    itemoption  = Trim(Left(itemoption,4))
    if (Len(itemoption)<4) then itemoption="0000"
    
    itemno      = ColData(10)
    itemcost    = ColData(11) 
    
    
    sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
    sqlStr = sqlStr + "itemoption, itemno, itemcost, itemvat, mileage, costtotal, orderdate," & vbCrlf
    sqlStr = sqlStr + "itemname,itemoptionname,makerid,buycash,buyvat,vatinclude,isupchebeasong,issailitem,oitemdiv)" & vbCrlf
    sqlStr = sqlStr + " select "
    sqlStr = sqlStr + " " + CStr(iid) + "" & vbCrlf
    sqlStr = sqlStr + " ,'" + orderserial + "'" & vbCrlf
    sqlStr = sqlStr + " ," + CStr(itemid) + "" & vbCrlf
    sqlStr = sqlStr + " ,'" + CStr(itemoption) + "'" & vbCrlf
    sqlStr = sqlStr + " ," + CStr(itemno) + "" & vbCrlf
    sqlStr = sqlStr + " ," + CStr(itemcost) + "" & vbCrlf
    sqlStr = sqlStr + " ,0,0,0,getdate()" & vbCrlf
    sqlStr = sqlStr + " ,i.itemname" & vbCrlf
    sqlStr = sqlStr + " ,IsNULL(o.optionname,'')" & vbCrlf
    sqlStr = sqlStr + " ,i.makerid" & vbCrlf
    sqlStr = sqlStr + " ,i.buycash" & vbCrlf
    sqlStr = sqlStr + " ,i.buyvat" & vbCrlf
    sqlStr = sqlStr + " ,i.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,(case when i.mwdiv='U' then 'Y' else 'N' end )" & vbCrlf
    sqlStr = sqlStr + " ,i.sailyn" & vbCrlf
    sqlStr = sqlStr + " ,i.itemdiv" & vbCrlf
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
    sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.itemoption='" + itemoption + "'"
    sqlStr = sqlStr + " where i.itemid=" + CStr(itemid)
    
    dbget.Execute sqlStr

next


''주문 master update
sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
sqlStr = sqlStr + " set totalvat = IsNULL(T.totalvat,0)," & vbCrlf
sqlStr = sqlStr + " totalcost = IsNULL(T.totalsum,0)," & vbCrlf
sqlStr = sqlStr + " totalsum = IsNULL(T.totalsum,0)," & vbCrlf
sqlStr = sqlStr + " subtotalprice = IsNULL(T.totalsum,0)" & vbCrlf
sqlStr = sqlStr + " from (" & vbCrlf
sqlStr = sqlStr + "     select sum(itemno*itemvat) as totalvat, sum(itemno*itemcost) as totalsum  " & vbCrlf
sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_detail " & vbCrlf
sqlStr = sqlStr + "     where orderserial='" + orderserial + "'"  & vbCrlf
sqlStr = sqlStr + " ) as T " & vbCrlf
sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial='" + orderserial + "'"

dbget.Execute sqlStr

''LimitedItemCheck orderserial,0,0,true

''현재고 업데이트
sqlStr = "exec [db_summary].[dbo].ten_RealtimeStock_regOrder '" & orderserial & "'"
dbget.execute sqlStr

''검토
dim addItemCount
sqlStr = " select count(orderserial) as cnt"
sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_detail"
sqlStr = sqlStr & " where orderserial='" & orderserial & "'"
sqlStr = sqlStr & " and itemid<>0"

rsget.Open sqlStr, dbget, 1
    addItemCount = rsget("cnt")
rsget.Close

if (rowcount<>addItemCount) then
	response.write "<script>alert('저장된 상품갯수 불일치 - 관리자 문의 요망." + CStr(rowcount) + "<>" + CStr(addItemCount) + "')</script>"
end if

dim referer
referer = request.ServerVariables("HTTP_REFERER")

response.write "<script>alert('저장 되었습니다. 주문번호 : " + orderserial + " ');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->