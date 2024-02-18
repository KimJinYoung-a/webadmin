<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, itemgubun, itemid, itemoption, stockReipgoDate
dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = RequestCheckvar(request("mode"),32)
itemgubun = RequestCheckvar(request("itemgubun"),2)
itemid = RequestCheckvar(request("itemid"),9)
itemoption = RequestCheckvar(request("itemoption"),4)
stockReipgoDate = RequestCheckvar(request("stockReipgoDate"),10)


if itemgubun="" then itemgubun="10"

dim sqlStr, ErrMsg

if (stockReipgoDate<>"") then
    on Error Resume next
    stockReipgoDate = Left(CStr(CDate(stockReipgoDate)),10)
    if Err then ErrMsg = Err.Description
    On Error Goto 0
end if

if (mode="stockreipgodate") then
    '' Stock 재입고일
    if (itemid<>"") or (itemoption<>"") or (ErrMsg<>"") then
        sqlStr = "exec [db_storage].[dbo].sp_Ten_StockReipgoSetting '" & itemgubun & "'," & itemid & ",'" & itemoption & "','" & stockReipgoDate & "'"	
        
        dbget.Execute sqlStr
    else    
        if (ErrMsg="") then ErrMsg = "Invalid Params - itemid"
    end if
    
elseif (mode="danjong") then
    '' 단종설정
    if (itemid<>"") then
        sqlStr = "update [db_item].[dbo].tbl_item"
        sqlStr = sqlStr + " set danjongyn='Y'"
        sqlStr = sqlStr + " ,lastupdate=getdate()" 
        sqlStr = sqlStr + " where itemid=" & CStr(itemid)
        
        dbget.Execute sqlStr
    else    
        ErrMsg = "Invalid Params - itemid"
    end if
    
elseif (mode="mssoldout") then
    '' MD품절
    if (itemid<>"") then
        sqlStr = "update [db_item].[dbo].tbl_item"
        sqlStr = sqlStr + " set danjongyn='M'"
        sqlStr = sqlStr + " ,lastupdate=getdate()" 
        sqlStr = sqlStr + " where itemid=" & CStr(itemid)
        
        dbget.Execute sqlStr
    else    
        ErrMsg = "Invalid Params - itemid"
    end if
    
else
    ErrMsg = "Invalid Params - Mode"
      
end if

%>

<script language="javascript">
<% if (ErrMsg<>"") then %>
    alert('<%= ErrMsg %>');
<% else %>
    opener.location.reload();
    alert('저장 되었습니다.');
<% end if %>
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->