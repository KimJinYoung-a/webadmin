<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
function GetCheckStatus(byVal sellsite, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp](sellsite, lastcheckdate, issuccess) "
	strSql = strSql + "		values('" & sellsite & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N') "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "

	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

Dim sellsite : sellsite = "gseshop"
Dim pdate : pdate = requestCheckVar(request("pdate"),10)
Dim act : act = requestCheckVar(request("act"),10)

Dim fromdate, selldate, todate, isSuccess

if (pdate<>"") and (act="touch") then
    Dim strSql
    strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " set lastcheckdate = '"&pdate&"', issuccess = 'Y' "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "
	dbget.Execute strSql
end if

Call GetCheckStatus(sellsite, selldate, isSuccess)

fromdate = selldate
todate = Left(Now, 10)

if (fromdate=pdate) and (pdate=todate) then
    fromdate = ""
elseif (fromdate=pdate) and (CDate(fromdate)<CDate(todate)) then
    fromdate = LEFT(CStr(dateadd("d",1,CDate(fromdate))),10)
end if

 
%>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
var g_thedate = "<%=fromdate%>";
$(function() {
    if (g_thedate==""){
         $("#idNoti").html("finish : <%=pdate%>")
        return;
    }
    $("#idNoti").html("start : "+g_thedate)
    var iurl = "http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=1003890&sdDt=<%=replace(fromdate,"-","")%>&tnsType=S"
    //var iurl = "http://webadmin.10x10.co.kr";
    setTimeout(function(){ 
        $("#idiframe").attr("src", iurl);  
    }, 1000);
    
});

function fnIfrmLoaded(){
    var isrc = $("#idiframe").attr("src");
    if (isrc!=""){
        // finish, touch date
        if (isrc.substring(0,23)=="http://ecb2b.gsshop.com"){
        //if (isrc.substring(0,22)=="http://webadmin.10x10.co.kr"){    
            $("#idNoti").html($("#idNoti").html()+"<br>finishd : "+g_thedate)
            alert("finished : "+g_thedate)
            location.href="?pdate="+g_thedate+"&act=touch";
        }else{
            alert(isrc.substring(0,23));
        }
    }
}

</script>

<table width="100%" align="left" cellpadding="3" cellspacing="0" class="table_tl">

	<tr height="25">
		<td class="td_br" colspan="2">
		<div id="idNoti"></div>	
		</td>
	</tr>
	<tr>
		<td class="td_br" colspan="2">
        <iframe onload="fnIfrmLoaded();" id="idiframe" name="idiframe" src="" width="100%" height="300" border="1"></iframe>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2" class="td_br">
		    <input type="button" class="button" value="´Ý±â" onClick="opener.location.reload();self.close();">
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->