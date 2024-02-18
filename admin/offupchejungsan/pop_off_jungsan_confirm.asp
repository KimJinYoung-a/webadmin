<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<% 
dim yyyymm : yyyymm = request("yyyymm")
dim mode   : mode   = request("mode") 
dim pyyyymm, yyyymmdd
dim sqlStr,resultRows
dim rdUrI

IF (mode="11") then
    pyyyymm = Left(dateadd("d",-1,dateadd("m",1,yyyymm+"-01")),10)
    
    rdUrI ="/admin/offshop/offshopjumun_error.asp?menupos=1183"
    rdUrI = rdUrI&"&yyyy1="&Left(yyyymm,4)&"&mm1="&Right(yyyymm,2)&"&dd1=01"
    rdUrI = rdUrI&"&yyyy2="&Left(pyyyymm,4)&"&mm2="&Mid(pyyyymm,6,2)&"&dd2="&right(pyyyymm,2)&""
    rdUrI = rdUrI&"&shopid="
    
    response.write "<script>location.replace('"&rdUrI&"');</script>"
    dbget.Close() : response.end
end if
    
IF (mode="1") then
    yyyymmdd = yyyymm + "-01"
    
    sqlStr = " select top 100 T.shopid,T.makerid"
    sqlStr = sqlStr& " ,sum(sellCnt) as sellCnt, sum(brandipchul) as brandipchul, sum(logicsipchul) as logicsipchul"
    sqlStr = sqlStr& " ,D.comm_cd,C.comm_name"
    sqlStr = sqlStr& " from ("
    sqlStr = sqlStr& " 	select  shopid,makerid,sum(d.itemno) as sellCnt, 0 as brandipchul, 0 as logicsipchul"
    sqlStr = sqlStr& " 	from db_shop.dbo.tbl_shopjumun_master m"
    sqlStr = sqlStr& " 		Join db_shop.dbo.tbl_shopjumun_detail d"
    sqlStr = sqlStr& " 		on m.orderno=d.orderno"
    sqlStr = sqlStr& " 	where m.shopregdate>='"&yyyymmdd&"'"
    sqlStr = sqlStr& " 	and m.cancelyn='N'"
    sqlStr = sqlStr& " 	and d.cancelyn='N'"
    sqlStr = sqlStr& " 	and isNULL(d.jcomm_cd,'')<>'B000'"
    sqlStr = sqlStr& " 	group by shopid,makerid"
    sqlStr = sqlStr& " 	union"
    sqlStr = sqlStr& " 	select m.shopid,d.designerid,0 as sellCnt, sum(d.itemno) as brandipchul, 0 as logicsipchul"
    sqlStr = sqlStr& " 	from db_shop.dbo.tbl_shop_ipchul_master m"
    sqlStr = sqlStr& " 		Join db_shop.dbo.tbl_shop_ipchul_detail d"
    sqlStr = sqlStr& " 		on m.idx=d.masteridx"
    sqlStr = sqlStr& " 		and m.execdt>='"&yyyymmdd&"'"
    sqlStr = sqlStr& " 	where m.deleteyn='N'"
    sqlStr = sqlStr& " 	and d.deleteyn='N'"
    sqlStr = sqlStr& " 	group by m.shopid,d.designerid"
    sqlStr = sqlStr& " 	union"
    sqlStr = sqlStr& " 	select m.socid,d.imakerid,0 as sellCnt, 0 as brandipchul, sum(d.itemno) as logicsipchul"
    sqlStr = sqlStr& " 	from db_storage.dbo.tbl_acount_storage_master m"
    sqlStr = sqlStr& " 		Join db_storage.dbo.tbl_acount_storage_detail d"
    sqlStr = sqlStr& " 		on m.code=d.mastercode"
    sqlStr = sqlStr& " 		and m.executedt>='"&yyyymmdd&"'"
    sqlStr = sqlStr& " 	where m.ipchulflag='S'"
    sqlStr = sqlStr& " 	and m.deldt is NULL"
    sqlStr = sqlStr& " 	and d.deldt is NULL"
    sqlStr = sqlStr& " 	group by m.socid,d.imakerid"
    sqlStr = sqlStr& " ) T"
    sqlStr = sqlStr& " 	left join db_summary.dbo.tbl_monthly_shop_designer SD"
    sqlStr = sqlStr& " 	on SD.yyyymm='"&yyyymm&"'"
    sqlStr = sqlStr& " 	and T.shopid=SD.shopid"
    sqlStr = sqlStr& " 	and T.makerid=SD.makerid"
    sqlStr = sqlStr& " 	left join db_shop.dbo.tbl_shop_designer D"
    sqlStr = sqlStr& " 	on T.shopid=D.shopid"
    sqlStr = sqlStr& " 	and T.makerid=D.makerid"
    sqlStr = sqlStr& " 	left join db_jungsan.dbo.tbl_jungsan_comm_code C"
	sqlStr = sqlStr& " 	on D.comm_cd=C.comm_cd"
    sqlStr = sqlStr& " where SD.shopid is NULL"
    sqlStr = sqlStr& " group by T.shopid,T.makerid,D.comm_cd,C.comm_name"
''rw sqlStr
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then 
        resultRows = rsget.getRows()   
    end if 
    rsget.close
ELSEIF (mode="2") then
    pyyyymm = LEFT(CStr(dateAdd("m",-1,yyyymm+"-01")),10)
    sqlStr = " select top 100 * from  db_shop.dbo.tbl_shopjumun_master"
    sqlStr = sqlStr& " where datediff(d,shopregdate,regdate)>1"
    sqlStr = sqlStr& " and shopregdate>='"&pyyyymm&"'"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then 
        resultRows = rsget.getRows()   
    end if 
    rsget.close
ELSEIF (mode="90") then
    pyyyymm = LEFT(CStr(dateAdd("m",-2,yyyymm+"-01")),7) 
    
    sqlStr = " select top 100 d.orderno,d.itemgubun,d.itemid,d.itemoption,m.makerid,count(*) as CNT"
    sqlStr = sqlStr& " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
    sqlStr = sqlStr& "      Join [db_jungsan].[dbo].tbl_off_jungsan_detail d"
    sqlStr = sqlStr& "      on m.idx=d.masteridx"
    sqlStr = sqlStr& "      left join [db_shop].[dbo].tbl_shop_item i "
    sqlStr = sqlStr& "      on d.itemgubun=i.itemgubun"
    sqlStr = sqlStr& "      and d.itemid=i.shopitemid"
    sqlStr = sqlStr& "      and d.itemoption=i.itemoption"
    sqlStr = sqlStr& " where 1=1"
    sqlStr = sqlStr& " and m.yyyymm>='"&pyyyymm&"'"
    sqlStr = sqlStr& " group by d.orderno,d.itemgubun,d.itemid,d.itemoption,m.makerid"
    sqlStr = sqlStr& " having count(*)>1"
'rw sqlStr    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then 
        resultRows = rsget.getRows()   
    end if 
    rsget.close
ELSE
    
END IF


dim i,j,cnt, cnt2, colCnt
if IsArray(resultRows) then
    cnt = Ubound(resultRows,2)
    colCnt = Ubound(resultRows,1)
else
    cnt = 0
end if
%>
<script language='javascript'>
function popOffContract(shopid,makerid){
    var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + makerid,"popshopupcheinfo","width=700 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<table width="700" border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#777777>
<form name="frm" method="get">
<input type="hidden" name="yyyymm" value="<%=yyyymm %>">
<tr bgcolor="#DDDDFF">
	<td>검토 구분 </td>
	<td bgcolor="#FFFFFF">
    	<select name="mode">
    	<option value=""  >선택
    	<option value="1" <%=CHKIIF(mode="1","selected","") %> >월별브랜드정산구분
    	<option value="2" <%=CHKIIF(mode="2","selected","") %> >주문 늦게 올린내역
    	
    	<option value="90" <%=CHKIIF(mode="90","selected","") %> >중복정산검토
    	</select>
	</td>
	<td bgcolor="#FFFFFF"><input type="button" value="검토" onClick="document.frm.submit();"></td>
</tr>
</form>
</table>
<p>

<% if mode<>"" then %>
<table width="700"  border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#777777>

<% if IsArray(resultRows) then %>
<% if (mode="1") then %>
<tr bgcolor="#DDDDFF">
    <td>매장ID</td>
    <td>브랜드ID</td>
    <td>판매량</td>
    <td>브랜드입출고</td>
    <td>물류입출고</td>
    <td>(현)정산구분</td>
    <td>(현)정산구분</td>
    <td></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
    <td colspan="30"><%'= sqlStr %></td>
</tr>
<% end if %>

<% if (mode="1") then %>
    <% for i=0 to cnt %>
    <tr bgcolor="#FFFFFF">
        <td><%= resultRows(0,i) %></td>
        <td><a href="javascript:popOffContract('<%= resultRows(0,i) %>','<%= resultRows(1,i) %>')"><%= resultRows(1,i) %></a></td>
        <td><%= resultRows(2,i) %></td>
        <td><%= resultRows(3,i) %></td>
        <td><%= resultRows(4,i) %></td>
        <td><%= resultRows(5,i) %></td>
        <td><%= resultRows(6,i) %></td>
        <td>&nbsp;</td>
    </tr>
    <% next %>
<% else %>
    <% for i=0 to cnt %>
    <tr bgcolor="#FFFFFF">
        <% for j=0 to colCnt %>
        <td><%= resultRows(j,i) %></td>
        <% next %>
    </tr>
    <% next %>
    
<% end if %>
<% else %>
<tr bgcolor="#FFFFFF">
    <td align="center">검색 내역이 없습니다.</td>
</tr>
<% end if %>
</table>
<% end if %>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->