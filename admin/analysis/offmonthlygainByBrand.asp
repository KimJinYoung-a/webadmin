<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/offReportClass.asp"-->
<%

Sub DrawOffJungsanGubun(selectBoxName,selectedId)
    dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select comm_cd, comm_name from [db_jungsan].[dbo].tbl_jungsan_comm_code"
   query1 = query1 + " where comm_group='Z002'"
   query1 = query1 + " and comm_isDel='N'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("comm_cd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("comm_cd")&"' "&tmp_str&">" + rsget("comm_cd") + " [" + rsget("comm_name") + "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   
   response.write("<option value='0000' "& chkIIF(selectedId="0000","selected","") &">     [미지정]</option>")
   response.write("</select>")
   
End Sub


dim yyyy1,mm1,dt
dim shopid, jungsangubun, makerid
yyyy1   = request("yyyy1")
mm1     = request("mm1")
shopid  = request("shopid")
jungsangubun = request("jungsangubun")
makerid = request("makerid")

if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim oReport
set oReport = New COffReport
oReport.FRectYYYYMM = yyyy1 & "-" & mm1
oReport.FRectShopid = shopid
oReport.FRectJungsanGubun = jungsangubun
oReport.FRectMakerid = makerid
oReport.GetShopMeachulByBrandJungsanGubun2


dim i, TTLitemcount, TTLRealSellSum, TTLjungsanitemcount, TTLjungsansum
%>
<script language='javascript'>
function reSearchByMakerid(makerid){
    var frm = document.frm;
    frm.makerid.value = makerid;
    frm.submit();
}


function PopJungSanDetailList(yyyymm,gubuncd,shopid){
    var popwin = window.open('redJungsandetail.asp?yyyymm=' + yyyymm + '&gubuncd=' + gubuncd + '&shopid=' + shopid,'redJungsandetail','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
    
    // "/admin/offupchejungsan/off_jungsandetailsum.asp?idx=39773&gubuncd=B012&shopid=streetshop011"
}


function popBrandMeachulDetailList(yyyymm,shopid,makerid){
    var yyyy1 = yyyymm.substr(0,4);
    var mm1 = yyyymm.substr(5,2);
    var dd1 = "01";
    
    var oDate = new Date(yyyy1, mm1, 0)
    
    var yyyy2 = oDate.getYear();
    var mm2 = oDate.getMonth()*1 + 1;
    var dd2 = oDate.getDate();
    
    mm2 = mm2.toString();
    if (mm2.length<2) mm2 = "0" + mm2;
    
    
    var params = "?menupos=452&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2;
    params = params + "&shopid=" + shopid + "&makerid=" + makerid;
    
    var popwin = window.open('/admin/offshop/itemsellsum.asp' + params,'shopMeachuldetail','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
    
    // "/admin/offshop/itemsellsum.asp?showtype=showtype&menupos=452&yyyy1=2007&mm1=12&dd1=25&yyyy2=2007&mm2=12&dd2=26&shopid=streetshop011&makerid=7321&rectorder=bysum&offgubun="
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간:<% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
		샾 : <% drawSelectBoxOffShop "shopid",shopid %>
		정산구분 : <% DrawOffJungsanGubun "jungsangubun", jungsangubun %>
		<br>
		브랜드 : <% DrawSelectBoxDesignerwithName "makerid",makerid %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a" >
<tr align="center" bgcolor="#E6E6E6" >
    <td width="90">매장</td>
    <td width="90">정산구분</td>
    <td width="140">브랜드ID</td>
    <td width="70">건수(상품)</td>
    <td width="90">매출</td>
    <td width="2"></td>
    <td width="70">건수(상품)</td>
    <td width="90">업체정산액</td>
    <td width="2"></td>
    <td width="70">건수(상품)</td>
    <td width="90">Center입고액</td>
    <td width="2"></td>
    <td width="70">건수(상품)</td>
    <td width="90">업체입고액</td>
    <td></td>
</tr>
<% for i=0 to oReport.FResultCount -1 %>
<%
TTLitemcount = TTLitemcount + oReport.FItemList(i).Ftotalitemcount
TTLRealSellSum = TTLRealSellSum + oReport.FItemList(i).FtotalRealSellSum
TTLjungsanitemcount = TTLjungsanitemcount + oReport.FItemList(i).Ftotaljungsanitemcount
TTLjungsansum       = TTLjungsansum + oReport.FItemList(i).Ftotaljungsansum

%>
<tr bgcolor="#FFFFFF"  align="center">
    <td><%= oReport.FItemList(i).FShopid %></td>
    <td><%= oReport.FItemList(i).FJungsanGubunName %></td>
    <td align="left"><a href="javascript:reSearchByMakerid('<%= oReport.FItemList(i).FMakerid %>');"><%= oReport.FItemList(i).FMakerid %></a></td>
    <td><%= FormatNumber(oReport.FItemList(i).Ftotalitemcount,0) %></td>
    <td align="right"><a href="javascript:popBrandMeachulDetailList('<%= yyyy1 %>-<%= mm1 %>','<%= oReport.FItemList(i).FShopid %>','<%= oReport.FItemList(i).Fmakerid %>')"><%= FormatNumber(oReport.FItemList(i).FtotalRealSellSum,0) %></a></td>
    <td></td>
    <td><%= FormatNumber(oReport.FItemList(i).FtotalJungsanitemcount,0) %></td>
    <td align="right"><a href="javascript:PopJungSanDetailList('<%= yyyy1 %>-<%= mm1 %>','<%= oReport.FItemList(i).FJungsanGubun %>','<%= oReport.FItemList(i).FShopid %>');"><%= FormatNumber(oReport.FItemList(i).FtotalJungsanSum,0) %></a></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF"  align="center">
    <td>합계</td>
    <td></td>
    <td></td>
    <td><%= FormatNumber(TTLitemcount,0) %></td>
    <td align="right"><%= FormatNumber(TTLRealSellSum,0) %></td>
    <td></td>
    <td><%= FormatNumber(TTLjungsanitemcount,0) %></td>
    <td align="right"><%= FormatNumber(TTLjungsansum,0) %></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
