<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
Dim makerid,yyyymm,shopid,mode,comm_cd,defaultmargin,defaultSuplymargin
makerid = requestCheckvar(request("makerid"),32)
yyyymm  = requestCheckvar(request("yyyymm"),32)
shopid  = requestCheckvar(request("shopid"),32)
mode    = requestCheckvar(request("mode"),32)
comm_cd         = requestCheckvar(request("comm_cd"),4)
defaultmargin   = requestCheckvar(request("defaultmargin"),10)
defaultSuplymargin  = requestCheckvar(request("defaultSuplymargin"),10)

dim ppyyyymm : ppyyyymm = Left(dateAdd("m",-2,yyyymm+"-01"),7)
dim nnyyyymm : nnyyyymm = Left(dateAdd("m",+2,yyyymm+"-01"),7)

Dim sqlStr, ArrList, i, AssignedRow

IF (mode="act") then
    sqlStr = " IF Exists(select * from db_summary.dbo.tbl_monthly_shop_designer where shopid='"&shopid&"' and makerid='"&makerid&"' and yyyymm='"&yyyymm&"')"& VbCRLF
    sqlStr = sqlStr & " BEGIN"& VbCRLF
    sqlStr = sqlStr & "     update db_summary.dbo.tbl_monthly_shop_designer"& VbCRLF
    sqlStr = sqlStr & "     set comm_cd='"&comm_cd&"'"& VbCRLF
    sqlStr = sqlStr & "     ,defaultmargin="&defaultmargin& VbCRLF
    sqlStr = sqlStr & "     ,defaultSuplymargin="&defaultSuplymargin& VbCRLF
    sqlStr = sqlStr & "     where shopid='"&shopid&"' and makerid='"&makerid&"' and yyyymm='"&yyyymm&"'"& VbCRLF
    sqlStr = sqlStr & " END"& VbCRLF
    sqlStr = sqlStr & " ELSE"& VbCRLF
    sqlStr = sqlStr & " BEGIN"& VbCRLF
    sqlStr = sqlStr & "     Insert Into db_summary.dbo.tbl_monthly_shop_designer"& VbCRLF
    sqlStr = sqlStr & "     (yyyymm,shopid,makerid,comm_cd,defaultmargin,defaultSuplymargin)"& VbCRLF
    sqlStr = sqlStr & "     values('"&yyyymm&"'"& VbCRLF
    sqlStr = sqlStr & "     ,'"&shopid&"'"& VbCRLF
    sqlStr = sqlStr & "     ,'"&makerid&"'"& VbCRLF
    sqlStr = sqlStr & "     ,'"&comm_cd&"'"& VbCRLF
    sqlStr = sqlStr & "     ,"&defaultmargin&""& VbCRLF
    sqlStr = sqlStr & "     ,"&defaultSuplymargin&""& VbCRLF
    sqlStr = sqlStr & "     )"& VbCRLF
    sqlStr = sqlStr & " END"& VbCRLF

    dbget.Execute sqlStr,AssignedRow

    IF (AssignedRow>0) then
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

sqlStr = " select top 10 D.YYYYMM,D.shopid,D.makerid,D.comm_cd,D.defaultmargin,D.defaultSuplymargin,C.comm_name"
sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_shop_designer D"
sqlStr = sqlStr & "     left join db_jungsan.dbo.tbl_jungsan_comm_code C"
sqlStr = sqlStr & "     on D.comm_cd=C.comm_cd"
sqlStr = sqlStr & " where shopid='"&shopid&"'"
sqlStr = sqlStr & " and makerid='"&makerid&"'"
sqlStr = sqlStr & " and yyyymm>='"&ppyyyymm&"'"
sqlStr = sqlStr & " and yyyymm<='"&nnyyyymm&"'"
sqlStr = sqlStr & " order by D.YYYYMM desc"

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	ArrList = rsget.getRows
end if
rsget.Close

dim cnt
If IsArray(ArrList) then
    cnt = UBound(ArrList,2)+1
ELSE
    cnt = 0
ENd IF

Dim OArrList
sqlStr = " select D.shopid,D.makerid,D.comm_cd,D.defaultmargin,D.defaultSuplymargin,C.comm_name"
sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer D"
sqlStr = sqlStr & "     left join db_jungsan.dbo.tbl_jungsan_comm_code C"
sqlStr = sqlStr & "     on D.comm_cd=C.comm_cd"
sqlStr = sqlStr & " where shopid='"&shopid&"'"
sqlStr = sqlStr & " and makerid='"&makerid&"'"
rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	OArrList = rsget.getRows
end if
rsget.Close

Dim pExists
%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;
    if (frm.comm_cd.value.length<1){
        alert('정산 구분을 선택 하세요.');
        frm.comm_cd.focus();
        return;
    }

    if (frm.defaultmargin.value.length<1){
        alert('매입 마진을 입력 하세요.');
        frm.defaultmargin.focus();
        return;
    }

    if (frm.defaultSuplymargin.value.length<1){
        alert('매장 공급 마진을 입력 하세요.');
        frm.defaultSuplymargin.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.mode.value="act";
        frm.submit();
    }
}

function assignThis(icomm_cd,idefaultmargin,idefaultSuplymargin){
    var frm = document.frmAct;
    frm.comm_cd.value=icomm_cd;
    frm.defaultmargin.value=idefaultmargin;
    frm.defaultSuplymargin.value=idefaultSuplymargin;

}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="50">현재기준</td>
	<td width="80">매장ID</td>
	<td width="140">브랜드ID</td>
	<td width="60">정산구분</td>
	<td width="60">매입마진</td>
	<td width="60">공급마진</td>
	<td width="60">비고</td>
</tr>
<% if IsArray(OArrList) then %>
<tr bgcolor="#FFFFFF" height="20">
    <td> 현재 </td>
    <td><%= OArrList(0,0) %></td>
    <td><%= OArrList(1,0) %></td>
    <td><%= OArrList(5,0) %></td>
    <td><%= OArrList(3,0) %></td>
    <td><%= OArrList(4,0) %></td>
    <td><input type="button" class="button" value="복사" onclick="assignThis('<%= OArrList(2,i) %>','<%= OArrList(3,i) %>','<%= OArrList(4,i) %>');"></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height="20">
    <td align="center" colspan="7"> 없음 </td>
</tr>
<% end if %>
</table>
<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="yyyymm" value="<%= yyyymm %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="50">년/월</td>
	<td width="80">매장ID</td>
	<td width="140">브랜드ID</td>
	<td width="60">정산구분</td>
	<td width="60">매입마진</td>
	<td width="60">공급마진</td>
	<td width="60">비고</td>
</tr>
<% for i=0 to cnt -1 %>
<tr bgcolor="#FFFFFF" height="20">
    <% if yyyymm=ArrList(1,i) then %>
    <% pExists = true %>
    <td><%= ArrList(0,i) %></td>
    <td><%= ArrList(1,i)  %></td>
    <td><%= ArrList(2,i) %></td>
    <td><% drawSelectBoxOFFJungsanCommCD "comm_cd",ArrList(3,i) %></td>
    <td><input type="text" name="defaultmargin" value="<%= ArrList(4,i) %>" size="4" maxlength="5"></td>
    <td><input type="text" name="defaultSuplymargin" value="<%= ArrList(5,i) %>" size="4" maxlength="5"></td>
    <td></td>
    <% else %>
    <td><%= ArrList(0,i) %></td>
    <td><%= ArrList(1,i) %></td>
    <td><%= ArrList(2,i) %></td>
    <td><%= ArrList(6,i) %></td>
    <td><%= ArrList(4,i) %></td>
    <td><%= ArrList(5,i) %></td>
    <td><input type="button" class="button" value="복사" onclick="assignThis('<%= ArrList(3,i) %>','<%= ArrList(4,i) %>','<%= ArrList(5,i) %>');"></td>
    <% end if %>
</tr>
<% next %>
<% if (Not pExists) then %>
<tr bgcolor="#FFF3F3" height="20">
    <td><%= yyyymm %></td>
    <td><%= shopid %></td>
    <td><%= makerid %></td>
    <td><% drawSelectBoxOFFJungsanCommCD "comm_cd","" %></td>
    <td><input type="text" name="defaultmargin" value="" size="4" maxlength="5"></td>
    <td><input type="text" name="defaultSuplymargin" size="4" maxlength="5"></td>
    <td></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="저장" onClick="saveThis()">
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
