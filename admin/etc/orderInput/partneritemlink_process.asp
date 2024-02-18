<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim mode,mallid,itemid
Dim outmallitemid, outmallitemname, outmallPrice, outmallSellYn, itemoption,P_itemoption, outmallitemOptionname
	mode = requestCheckVar(request("mode"),16)
	mallid = requestCheckVar(request("mallid"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemoption = requestCheckVar(request("itemoption"),4)
	p_itemoption = requestCheckVar(request("p_itemoption"),4)
	outmallitemid = requestCheckVar(request("outmallitemid"),32)
	outmallitemname = requestCheckVar(request("outmallitemname"),100)
	outmallPrice = requestCheckVar(request("outmallPrice"),10)
	outmallSellYn = requestCheckVar(request("outmallSellYn"),10)
	outmallitemOptionname= requestCheckVar(request("outmallitemOptionname"),100)

outmallPrice = replace(outmallPrice,",","")
outmallitemname = Trim(outmallitemname)

response.write "mode:"&mode&"<Br>"

if (mallid="") then 
    rw "Require mallid"
    response.end    
end if

Dim sqlStr, itemExists, AssignedRow
Dim iExists

iExists = false
''CHECK ITEM
'IF (application("Svr_Info")	<> "Dev") then

sqlStr = "select count(*) as CNT from db_item.dbo.tbl_item where itemid="&itemid

rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    iExists = rsget("CNT")>0
	end if
rsget.close

if (Not iExists) then
    response.write "<script>alert('"&itemid&" 상품코드가 존재하지 않습니다.');history.back()</script>"
    response.end    
end if

iExists = false
if (itemoption<>"") then
    if (itemoption<>"0000") then
        sqlStr = "select count(*) as CNT from db_item.dbo.tbl_item_option where itemid="&itemid&" and itemoption='"&itemoption&"'"
       ''rw sqlStr
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            iExists = rsget("CNT")>0
        end if
        rsget.close
        
        if (Not iExists) then
            ''response.write "<script>alert('"&itemid&" 옵션코드가 존재하지 않습니다. 옵션이 없는 경우 또는 옵션별 매칭이 필요한 경우만 입력');history.back()</script>"
            rw itemid&" 옵션코드가 존재하지 않습니다. 옵션이 없는 경우 또는 옵션별 매칭이 필요한 경우만 입력"
            response.end    
        end if
    else
        sqlStr = "select count(*) as CNT from db_item.dbo.tbl_item_option where itemid="&itemid
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            iExists = rsget("CNT")>0
        end if
        rsget.close
        
        if (iExists) then
            ''response.write "<script>alert('"&itemid&" 옵션이 존재하는 상품입니다. 0000 입력 불가');history.back()</script>"
            rw itemid&" 옵션이 존재하는 상품입니다. 0000 입력 불가"
            response.end    
        end if
        
    end if
end if

if (mode="add") then
    sqlStr = "select count(*) as CNT"
    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_EtcItemLink"
    sqlStr = sqlStr & " where mallid='"&mallid&"'"
    
    '/제휴상품명과 제휴옵션명으로 매칭하는 제휴몰
    if GetItemMaeching_itemname_itemoptionname(mallid) then
    	sqlStr = sqlStr & " and outmallitemname='"&outmallitemname&"'"
    	sqlStr = sqlStr & " and outmallitemOptionname='"&outmallitemOptionname&"'"
    else
    	sqlStr = sqlStr & " and outmallitemid='"&outmallitemid&"'"
    	sqlStr = sqlStr & " and outmallitemOptionname='"&outmallitemOptionname&"'"
    end if
    
    'response.write sqlStr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        iExists = rsget("CNT")>0
    end if
    rsget.close
    
    if (iExists) then
        response.write "<script>alert('"&outmallitemid&" ["&outmallitemOptionname&"] 이미 등록된 상품코드 [옵션명] 입니다.');history.back()</script>"
        response.end    
    end if
end if
'end if

sqlStr = "select * from  db_temp.dbo.tbl_xSite_EtcItemLink"
sqlStr = sqlStr & " where mallid='"&mallid&"' and itemid="&itemid&" and itemoption='"&p_itemoption&"'"

'response.write sqlStr & "<Br>"
rsget.Open sqlStr,dbget,1

if rsget.Eof then
    itemExists = FALSE
ELSE
    itemExists = TRUE
end if

rsget.close

if (itemExists) then
    IF (mode="del") then
        sqlStr = "delete from db_temp.dbo.tbl_xSite_EtcItemLink" & VBCRLF
        sqlStr = sqlStr & " where mallid='"&mallid&"' and itemid="&itemid & VBCRLF
        sqlStr = sqlStr & " and itemoption='"&p_itemoption&"'"
        dbget.Execute sqlStr,AssignedRow
    ELSE
        sqlStr = "Update db_temp.dbo.tbl_xSite_EtcItemLink" & VBCRLF
        sqlStr = sqlStr & " SET outmallPrice="&outmallPrice & VBCRLF
        IF outmallitemid="" then
            sqlStr = sqlStr & " ,outmallitemid=NULL"
        ELSE
            sqlStr = sqlStr & " ,outmallitemid='"&outmallitemid&"'" & VBCRLF
        ENd IF
        IF outmallitemname="" then
            sqlStr = sqlStr & " ,outmallitemname=NULL" & VBCRLF
        ELSE
            sqlStr = sqlStr & " ,outmallitemname='"&html2DB(outmallitemname)&"'" & VBCRLF
        ENd If
        sqlStr = sqlStr & " ,outmallitemOptionname='"&html2DB(outmallitemOptionname)&"'" & VBCRLF
        sqlStr = sqlStr & " ,outmallSellYn='"&outmallSellYn&"'" & VBCRLF
        sqlStr = sqlStr & " , itemoption='"&itemoption&"'"
        sqlStr = sqlStr & " where mallid='"&mallid&"' and itemid="&itemid & VBCRLF
        sqlStr = sqlStr & " and itemoption='"&p_itemoption&"'"
        dbget.Execute sqlStr,AssignedRow
    end if
else
    sqlStr = "Insert Into db_temp.dbo.tbl_xSite_EtcItemLink"
    sqlStr = sqlStr & " (itemid,itemoption,mallID,outmallitemid,outmallitemname,outmallitemOptionname,outmallPrice,outmallSellYn)"
    sqlStr = sqlStr & " values("
    sqlStr = sqlStr & " "&itemid&VbCRLF
    sqlStr = sqlStr & " ,'"&itemoption&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&outmallitemid&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&html2DB(outmallitemname)&"'"&VbCRLF
    sqlStr = sqlStr & " ,'"&html2DB(outmallitemOptionname)&"'"&VbCRLF
    sqlStr = sqlStr & " ,"&(outmallPrice)&""&VbCRLF
    sqlStr = sqlStr & " ,'"&outmallSellYn&"'"&VbCRLF
    sqlStr = sqlStr & " )"
    dbget.Execute sqlStr,AssignedRow
end if
%>

<script language='javascript'>
	<% if (mode="del") then %>
		alert('<%=AssignedRow %>건 삭제됨.')
		opener.location.reload();
		window.close();
	<% else %>
		alert('<%=AssignedRow %>건 반영됨')
		location.href="/admin/etc/orderInput/partneritemlink_modify.asp?mallid=<%=mallid %>&itemid=<%=itemid%>&itemoption=<%=itemoption%>"
	<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->