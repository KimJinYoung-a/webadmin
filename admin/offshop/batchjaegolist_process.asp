<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, jobshopid, jobgubun, idx, jobkey, newIdx
dim buf, i
dim sqlStr
	mode        = requestCheckVar(request.Form("mode"),32)
	jobshopid   = requestCheckVar(request.Form("jobshopid"),32)
	jobgubun    = requestCheckVar(request.Form("jobgubun"),10)
	idx         = request.Form("idx")

''091116 011 04 38864 : orderNo Style
dim TmpOrderNo : TmpOrderNo = Right(Replace(Left(CStr(now()),10),"-",""),6) & "000" & "00" 
dim InvalidExists : InvalidExists = false

jobkey = 0

if mode="arr" then
    response.write "idx="&idx&"<br>"
	idx = split(idx,"|")

	for i = 0 to Ubound(idx) - 1
        if (trim(idx(i)) <> "") then
            if (buf = "") then
                    buf = requestCheckVar(trim(idx(i)),10)
            else
                    buf = buf + "," + requestCheckVar(trim(idx(i)),10)
            end if
        end if
	next
    response.write "buf="&buf&"<br>"
    
    sqlStr = "select count(*) as CNT from [db_shop].[dbo].tbl_shop_tempstock_master "
    sqlStr = sqlStr + " where idx in (" + buf + ") "
    sqlStr = sqlStr + " and IsNULL(jobstate,0)>0"
    rsget.Open sqlStr,dbget,1
        InvalidExists = rsget("CNT")>0
    rsget.Close
    
    if (InvalidExists) then
        response.write "<script type='text/javascript'>alert('미처리 상태만 작업일괄지정이 가능합니다.');</script> "
        response.write "<script type='text/javascript'>history.back();</script>"
        response.end
    end if
    
    
    sqlStr = "select max(isnull(jobkey, 0)) as jobkey from [db_shop].[dbo].tbl_shop_tempstock_master "
    rsget.Open sqlStr,dbget,1
	    jobkey = rsget("jobkey") + 1
	rsget.Close
    
    dbget.BeginTrans
    
    sqlStr = " insert into [db_shop].[dbo].tbl_shop_tempstock_master " & VbCrlf
    sqlStr = sqlStr & " (orderno,shopid,totalsum,realsum,jumundiv,jumunmethod " & VbCrlf
    sqlStr = sqlStr & " ,shopregdate,cancelyn " & VbCrlf
    sqlStr = sqlStr & " ,shopidx,jobshopid, jobgubun, jobkey,subjobkey,jobstate) " & VbCrlf
    sqlStr = sqlStr & " select '' as orderno " & VbCrlf
    sqlStr = sqlStr & " ,Max(shopid) as shopid " & VbCrlf
    sqlStr = sqlStr & " ,sum(totalsum) as totalsum " & VbCrlf
    sqlStr = sqlStr & " ,sum(realsum) as realsum " & VbCrlf
    sqlStr = sqlStr & " ,'00' as jumundiv " & VbCrlf
    sqlStr = sqlStr & " ,'90' as jumunmethod " & VbCrlf
    sqlStr = sqlStr & " ,Max(shopregdate) as shopregdate " & VbCrlf
    sqlStr = sqlStr & " ,'N' as cancelyn " & VbCrlf
    sqlStr = sqlStr & " ,0 as shopidx " & VbCrlf
    sqlStr = sqlStr & " ,'" + CStr(jobshopid) + "' as jobshopid " & VbCrlf
    sqlStr = sqlStr & " ,'10' as jobgubun " & VbCrlf
    sqlStr = sqlStr & " ," + CStr(jobkey) + " as jobkey " & VbCrlf
    sqlStr = sqlStr & " ,0 as subjobkey " & VbCrlf
    sqlStr = sqlStr & " ,3 as jobstate " & VbCrlf
    sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_tempstock_master " & VbCrlf
    sqlStr = sqlStr & " where idx in (" + buf + ") " & VbCrlf
    sqlStr = sqlStr & " and cancelyn='N' " & VbCrlf
    sqlStr = sqlStr & " and IsNULL(jobstate,0)=0 " & VbCrlf
    dbget.Execute sqlStr
    
    sqlStr = " select @@Identity as idt"
''response.write sqlStr   
    rsget.Open sqlStr,dbget,1
        newIdx = rsget("idt")
    rsget.Close
    
    TmpOrderNo = TmpOrderNo & Format00(5,newIdx)
    
    sqlStr = "update [db_shop].[dbo].tbl_shop_tempstock_master " & VbCrlf
    sqlStr = sqlStr + " set orderno='" & TmpOrderNo & "'"  & VbCrlf
    sqlStr = sqlStr + " ,jobshopid = '" + CStr(jobshopid) + "'" & VbCrlf
    sqlStr = sqlStr + " ,jobstate = 3" & VbCrlf
    sqlStr = sqlStr + " ,jobgubun = '" + jobgubun + "'" & VbCrlf
    sqlStr = sqlStr + " ,jobkey = " + CStr(jobkey) + " " & VbCrlf
    sqlStr = sqlStr + " where idx=" & newIdx
    dbget.Execute sqlStr

    sqlStr = "insert into [db_shop].[dbo].tbl_shop_tempstock_detail" & VbCrlf
    sqlStr = sqlStr + " (" & VbCrlf
    sqlStr = sqlStr + " masteridx," & VbCrlf
    sqlStr = sqlStr + " orderno," & VbCrlf
    sqlStr = sqlStr + " itemgubun," & VbCrlf
    sqlStr = sqlStr + " itemid," & VbCrlf
    sqlStr = sqlStr + " itemoption," & VbCrlf
    sqlStr = sqlStr + " itemname," & VbCrlf
    sqlStr = sqlStr + " itemoptionname," & VbCrlf
    sqlStr = sqlStr + " sellprice," & VbCrlf
    sqlStr = sqlStr + " realsellprice," & VbCrlf
    sqlStr = sqlStr + " suplyprice," & VbCrlf
    sqlStr = sqlStr + " itemno," & VbCrlf
    sqlStr = sqlStr + " makerid," & VbCrlf
    sqlStr = sqlStr + " jungsanid," & VbCrlf
    sqlStr = sqlStr + " cancelyn," & VbCrlf
    sqlStr = sqlStr + " shopidx," & VbCrlf
    sqlStr = sqlStr + " discountprice," & VbCrlf
    sqlStr = sqlStr + " shopbuyprice" & VbCrlf
    sqlStr = sqlStr + " )" & VbCrlf
    sqlStr = sqlStr + " select " & CStr(newIdx) & " as masteridx, '" & TmpOrderNo & "' as tmpOrderNo" & VbCrlf
    sqlStr = sqlStr + " ,d.itemgubun, d.itemid, d.itemoption" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemname, s.shopitemoptionname" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemprice" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemprice" & VbCrlf
    sqlStr = sqlStr + " ,(Case When IsNULL(s.shopsuplycash,0)<>0 then s.shopsuplycash else convert(int,s.shopitemprice*(100-IsNULL(F.defaultmargin,35))/100) end) as shopsuplycash" & VbCrlf
    sqlStr = sqlStr + " ,sum(d.itemno) as itemno" & VbCrlf
    sqlStr = sqlStr + " ,s.makerid" & VbCrlf
    sqlStr = sqlStr + " ,s.makerid as jungsanid" & VbCrlf
    sqlStr = sqlStr + " ,'N' as cancelyn" & VbCrlf
    sqlStr = sqlStr + " ,0,0,0" & VbCrlf
    sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_detail d" & VbCrlf
    sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_item s" & VbCrlf   '''shop_item에 상품이 없으면 안됨..
    sqlStr = sqlStr + " 	on d.itemgubun=s.itemgubun " & VbCrlf
    sqlStr = sqlStr + " 	and d.itemid=s.shopitemid" & VbCrlf
    sqlStr = sqlStr + " 	and d.itemoption=s.itemoption" & VbCrlf
    sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer F"
    sqlStr = sqlStr + " 	on d.makerid=F.makerid"
    sqlStr = sqlStr + " 	and F.shopid='"&jobshopid&"'"
    sqlStr = sqlStr + " where d.masteridx in (" + buf + ")" & VbCrlf
    sqlStr = sqlStr + " and cancelyn='N'" & VbCrlf
    sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemname, s.shopitemoptionname" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemprice" & VbCrlf
    sqlStr = sqlStr + " ,s.shopitemprice" & VbCrlf
    sqlStr = sqlStr + " ,(Case When IsNULL(s.shopsuplycash,0)<>0 then s.shopsuplycash else convert(int,s.shopitemprice*(100-IsNULL(F.defaultmargin,35))/100) end)" & VbCrlf
    sqlStr = sqlStr + " ,s.makerid" & VbCrlf
    dbget.Execute sqlStr
    
    sqlStr = "update [db_shop].[dbo].tbl_shop_tempstock_master " & VbCrlf
    sqlStr = sqlStr + " set jobshopid = '" + CStr(jobshopid) + "'" & VbCrlf
    sqlStr = sqlStr + " , jobstate = 3" & VbCrlf
    sqlStr = sqlStr + " , jobgubun = '" + jobgubun + "'" & VbCrlf
    sqlStr = sqlStr + " , jobkey = " + CStr(jobkey) + " " & VbCrlf
    sqlStr = sqlStr + " , subjobkey = " + CStr(newIdx) + " " & VbCrlf
    sqlStr = sqlStr + " where idx in (" + buf + ") "
    dbget.Execute sqlStr
    'response.write sqlStr
    
    dbget.commitTrans
end if

if mode="cancelarr" then
	idx = split(idx,"|")

	for i = 0 to Ubound(idx) - 1
        if (trim(idx(i)) <> "") then
            if (buf = "") then
                    buf = requestCheckVar(trim(idx(i)),10)
            else
                    buf = buf + "," + requestCheckVar(trim(idx(i)),10)
            end if
        end if
	next
	
    sqlStr = "select count(*) as CNT from [db_shop].[dbo].tbl_shop_tempstock_master "
    sqlStr = sqlStr + " where idx in (" + buf + ") "
    sqlStr = sqlStr + " and jobstate<>3"
    rsget.Open sqlStr,dbget,1
        InvalidExists = rsget("CNT")>0
    rsget.Close
    
    if (InvalidExists) then
        response.write "<script type='text/javascript'>alert('처리중 상태가 아닙니다. 작업지정을 취소할 수 없습니다.');</script> "
        response.write "<script type='text/javascript'>history.back();</script>"
        response.end
    end if
    
    dbget.BeginTrans
    
    sqlStr = "delete from [db_shop].[dbo].tbl_shop_tempstock_master "
    sqlStr = sqlStr + " where idx in (" + buf + ") "
    dbget.Execute sqlStr
    
    sqlStr = "delete from [db_shop].[dbo].tbl_shop_tempstock_detail "
    sqlStr = sqlStr + " where masteridx in (" + buf + ") "
    dbget.Execute sqlStr
    
    sqlStr = "update [db_shop].[dbo].tbl_shop_tempstock_master "
    sqlStr = sqlStr + " set jobshopid = null, jobgubun = null, jobkey = null, jobstate=0, subjobkey=NULL "
    sqlStr = sqlStr + " where subjobkey in (" + buf + ") "
    dbget.Execute sqlStr
    
    dbget.commitTrans
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script type='text/javascript'>
	alert('수정되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->