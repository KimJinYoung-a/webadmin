<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%

function GetStoreCateName(icatecode)
    dim sqlStr
    sqlStr = "select top 1 storeCateName "
    sqlStr = sqlStr + " from [db_temp].dbo.tbl_interpark_Tmp_StoreCategory "
    sqlStr = sqlStr + " where StoreCateCode='" + icatecode + "'"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        GetStoreCateName = rsget("storeCateName")
    end if   
    rsget.close
end function

function GetDispCateName(icatecode)
    dim sqlStr
    sqlStr = "select top 1 DispCateName "
    sqlStr = sqlStr + " from [db_temp].dbo.tbl_interpark_Tmp_DispCategory "
    sqlStr = sqlStr + " where DispCateCode='" + icatecode + "'"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        GetDispCateName = rsget("DispCateName")
    end if   
    rsget.close
end function

dim itemid , mode, SupplyCtrtSeq, interparkstorecategory, interparkDispCategory
itemid  = RequestCheckVar(request("itemid"),9)
mode    = request("mode")
SupplyCtrtSeq = request("SupplyCtrtSeq")
interparkstorecategory = request("interparkstorecategory")
interparkDispCategory = request("interparkDispCategory")

dim sqlStr, AssignRow
if (mode="editSq") and (itemid<>"") and (SupplyCtrtSeq<>"") then
    sqlStr = "update [db_item].[dbo].tbl_interpark_reg_item" & VbCrlf
    sqlStr = sqlStr & " set interParkSupplyCtrtSeq=" & SupplyCtrtSeq & VbCrlf
    if (interparkstorecategory="") then
        sqlStr = sqlStr & " , interparkstorecategory=NULL"  & VbCrlf
    else
        sqlStr = sqlStr & " , interparkstorecategory='" & interparkstorecategory & "'" & VbCrlf
    end if
    
    if (interparkDispCategory="") then
        sqlStr = sqlStr & " , PinterParkDispCategory=NULL"  & VbCrlf
    else
         sqlStr = sqlStr & " , PinterParkDispCategory='" & interparkDispCategory & "'" & VbCrlf
    end if

    sqlStr = sqlStr & " where itemid=" & itemid

    dbget.Execute sqlStr,AssignRow
    
    if (AssignRow<1) then
        response.write "<script>alert('수정되지 않았습니다. 미등록상품의 경우 수정 불가');</script>"
    end if
end if

dim oInterParkitem, oSupplyCtrtSeq, oSupplyCtrtSeqName, ointerparkstorecategory, ointerparkstorecategoryTxt
dim ointerparkdispcategory, ointerparkdispcategoryTxt
set oInterParkitem = New CiParkRegItem
oInterParkitem.GetIParkOneItemList itemid, (mode="sellStatNONE")

if (oInterParkitem.FResultCount>0) then
    oSupplyCtrtSeq = oInterParkitem.FItemList(0).FSupplyCtrtSeq
    oSupplyCtrtSeqName = oInterParkitem.FItemList(0).getSupplyCtrtSeqName
    ointerparkstorecategory = oInterParkitem.FItemList(0).Finterparkstorecategory
    ointerparkdispcategory = oInterParkitem.FItemList(0).Finterparkdispcategory
end if

set oInterParkitem = Nothing

if (ointerparkstorecategory<>"") then
    ointerparkstorecategoryTxt = GetStoreCateName(ointerparkstorecategory)
end if

if (ointerparkdispcategory<>"") then
    ointerparkdispcategoryTxt = GetDispCateName(ointerparkdispcategory)
end if

if (oSupplyCtrtSeq="") or isNULL(oSupplyCtrtSeq) then
    oSupplyCtrtSeq = "2"
    oSupplyCtrtSeqName = "리빙"
end if
%>
<script language='javascript'>
function editSupplyCtrtSeq(frm){
    if (confirm('수정 하시겠습니까?')){
        frm.mode.value="editSq";
        frm.submit();
    }
}

function searchCate(frm){
    if (frm.sRect.value.length<1){
        //alert('검색어를 입력하세요.');
        //frm.sRect.focus();
        //return;
    }
    
    
    frm.action="/admin/etc/iframeInterParkDispcateSelect.asp"
    frm.target = "iFrameDispCate";
    frm.submit();
}

function searchStoreCate(frm){
    frm.action="/admin/etc/iframeInterParkStoreCateSelect.asp"
    frm.target = "iFrameStoreCate";
    frm.submit();
}

function getOnload(){
    alert('해당 상품의 전시 카테고리 매핑을 수정합니다.\n\n이속성은 카테고리별 매핑보다 우선적용됩니다.');
    
}
<% if NOT ((mode="editSq") and (itemid<>"") and (SupplyCtrtSeq<>"")) then %>
window.onload=getOnload;
<% end if %>

</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" class="a">
<tr>
    <td>
        <table width="280" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" class="a">
        <form name="frmSvr" method="post" action="">
        <input type="hidden" name="itemid" value="<%= itemid %>">
        <input type="hidden" name="mode" value="">
        <tr bgcolor="#FFFFFF">
            <td width="80" bgcolor="#F3F3FF">전시1</td> 
            <td>
                <input type="text" class="text" name="interparkdispcategory" value="<%= ointerparkdispcategory %>" size="32" maxlength="32">
                <input type="text" class="text_ro" name="interparkdispcategoryText" value="<%= ointerparkdispcategoryTxt %>" size="32" >
            </td> 
        </tr>
        <tr bgcolor="#FFFFFF">
            <td colspan="2" height="100"></td>
        </tr>
        
        <tr height="50">
            <td width="80">샵구분</td>
            <td align="left">
                <input type="text" class="text_ro" name="SupplyCtrtSeq" value="<%= oSupplyCtrtSeq %>" size="2"> 
                <input class="text_ro" type="text" name="SupplyCtrtSeqName" size="8" value="<%= oSupplyCtrtSeqName %>">
            </td>
        </tr>
        <% if (FALSE) then %>
        <tr bgcolor="#FFFFFF">
            <td width="80">카테고리</td>
            <td>
                <input type="text" class="text" name="interparkstorecategory" value="<%= ointerparkstorecategory %>" size="32" maxlength="32">
                <input type="text" class="text_ro" name="interparkstorecategoryText" value="<%= ointerparkstorecategoryTxt %>" size="32" >
            </td> 
        </tr>
        <% else %>
        <input type="hidden" name="interparkstorecategory" value="<%= ointerparkstorecategory %>" >
        <input type="hidden" name="interparkstorecategoryText" value="<%= ointerparkstorecategoryTxt %>" >
        <% end if %>
        <tr height="30">
            <td align="center" colspan="2" >
                <input type="button" value="수정" onClick="editSupplyCtrtSeq(frmSvr)">
            </td>
        </tr>
        </form>
        </table>
    </td>
    <td>
        <table width="400" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
        <form name="frmDispSearch" >
        <input type="hidden" name="mode" value="all">
        <tr  bgcolor="#FFFFFF">    
            <td>
            <input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchCate(frmDispSearch);" ><input type="button" class="button" value="검색" onClick="searchCate(frmDispSearch);">  
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td><iframe name="iFrameDispCate" id="iFrameDispCate" src="/admin/etc/iframeInterParkDispcateSelect.asp" width="600" height="180" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe></td>
        </tr>
        </form>
        <!-- 더이상 사용안함 
        <form name="frmStoreCateSearch" >
        <input type="hidden" name="mode" value="all">
        <tr  bgcolor="#FFFFFF">    
            <td>
            <input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchStoreCate(frmStoreCateSearch);" ><input type="button" class="button" value="검색" onClick="searchStoreCate(frmStoreCateSearch);">  
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td><iframe name="iFrameStoreCate" id="iFrameStoreCate" src="/admin/etc/iframeInterParkStoreCateSelect.asp" width="600" height="180" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe></td>
        </tr>
        </form>
        -->
        </table>
    </td>
    </tr>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->