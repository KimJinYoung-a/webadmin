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
dim shopid, jungsangubun
yyyy1   = request("yyyy1")
mm1     = request("mm1")
shopid  = request("shopid")
jungsangubun = request("jungsangubun")

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
oReport.GetShopMeachulByJungsanGubun2


dim i, TTLitemcount, TTLRealSellSum, TTLjungsanitemcount, TTLjungsansum
%>
<script language='javascript'>
function ReSearchMeachulByJungsanGubun(shopid,jungsangubun){
    var frm = document.frm;
    frm.shopid.value = shopid;
    
    if (jungsangubun=="    "){
        jungsangubun = "0000";
    }
    
    frm.jungsangubun.value = jungsangubun;
    
    frm.submit();
}

function popBrandMeachulByJungsanGubun(shopid,jungsangubun){
    if (jungsangubun=="    "){
        jungsangubun = "0000";
    }
    
    var popwin = window.open('offmonthlygainByBrand.asp?menupos=<%= menupos %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&shopid=' + shopid + '&jungsangubun=' + jungsangubun,'popoffmonthlygainByBrand','width=900,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF" class="a" >
<tr>
    <td>
        매출수익 <br>
         : 특정, 업체특정, 매장매입 = 매출 - 업체정산액 <br>
         : 매입, ETC = 매출 - 센터입고액<br>
         <br>
         오프매입은 매장 구분없이 일괄매입<br>
         텐매입 : Center 매입구분이 매입인것<br>
         출고매입 : 입출고 내역으로 매입정산 (특정, 매입, 업체배송 존재가능)
    </td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간:<% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
		샾 : <% drawSelectBoxOffShop "shopid",shopid %>
		정산구분 : <% DrawOffJungsanGubun "jungsangubun", jungsangubun %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a" >
<tr align="center" bgcolor="#E6E6E6" >
    <td width="90">매장</td>
    <td width="100">정산구분</td>
    <td width="70">건수(상품)</td>
    <td width="90">매출</td>
    <td width="1"></td>
    <td width="70">건수(상품)</td>
    <td width="90">업체정산액</td>
    <td width="1"></td>
    <td width="70">건수(상품)</td>
    <td width="90">Center입고액</td>
    <td width="1"></td>
    <td width="70">건수(상품)</td>
    <td width="90">업체입고액</td>
    <td width="1"></td>
    <td>매출수익</td>
</tr>
<% for i=0 to oReport.FResultCount -1 %>
<%
TTLitemcount        = TTLitemcount + oReport.FItemList(i).Ftotalitemcount
TTLRealSellSum      = TTLRealSellSum + oReport.FItemList(i).FtotalRealSellSum
TTLjungsanitemcount = TTLjungsanitemcount + oReport.FItemList(i).Ftotaljungsanitemcount
TTLjungsansum       = TTLjungsansum + oReport.FItemList(i).Ftotaljungsansum
%>
<tr bgcolor="#FFFFFF"  align="center">
    <td><a href="javascript:ReSearchMeachulByJungsanGubun('<%= oReport.FItemList(i).FShopid %>','');"><%= oReport.FItemList(i).FShopid %></a></td>
    <td><a href="javascript:ReSearchMeachulByJungsanGubun('','<%= oReport.FItemList(i).FJungsanGubun %>');"><%= oReport.FItemList(i).FJungsanGubunName %></a></td>
    <td><%= FormatNumber(oReport.FItemList(i).Ftotalitemcount,0) %></td>
    <td align="right"><a href="javascript:popBrandMeachulByJungsanGubun('<%= oReport.FItemList(i).FShopid %>','<%= oReport.FItemList(i).FJungsanGubun %>')"><%= FormatNumber(oReport.FItemList(i).FtotalRealSellSum,0) %></a></td>
    <td></td>
    <td><%= FormatNumber(oReport.FItemList(i).FtotalJungsanitemcount,0) %></td>
    <td align="right"><a href="javascript:popBrandMeachulByJungsanGubun('<%= oReport.FItemList(i).FShopid %>','<%= oReport.FItemList(i).FJungsanGubun %>')"><%= FormatNumber(oReport.FItemList(i).FtotalJungsanSum,0) %></a></td>
    <td></td>
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
    <td></td>
</tr>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
