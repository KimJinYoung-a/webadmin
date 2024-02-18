<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->

<%
dim mode, id, segumil
dim page
dim extsitename

mode = request("mode")
id   = request("id")
segumil = request("segumil")

dim sqlStr
if (mode="segumil") then
	sqlstr = "update [db_jungsan].dbo.tbl_etcsite_jungsanmaster" + VbCrlf
    sqlstr = sqlstr + " set segumil = '" + segumil + "'"  + VbCrlf
    sqlstr = sqlstr + " where id = " + CStr(id) + ""  + VbCrlf

    rsget.Open sqlstr, dbget, 1

end if

extsitename = request("extsitename")

page = request("page")
if (page="") then page=1



dim ijungsan
set ijungsan = new CUpcheJungSan


ijungsan.FcurrPage = page
ijungsan.FPageSize=40
ijungsan.getDefaultInfo extsitename

ijungsan.FrectSiteName = extsitename
ijungsan.PartnerOldJungSanDeasangList

dim ix
dim bufsum, deasangsum, amountsum
bufsum =0
deasangsum =0
amountsum =0
%>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
function ViewOrderDetail(frm){
    frm.target = 'orderdetail';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();
}

function jungsanopen(num){
  window.open("jungsan_window.asp?masterid=" + num ,"window","status=no,toolbar=no,resizable=no,scrollbars=no, menubar=no,width=300,height=200");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="900" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" id="sgDt" value="">
	<tr>
		<td class="a" width="500">
		사이트:
		<% drawSelectBoxPartner "extsitename",extsitename %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="900" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDDD">
	<td colspan="11">
	<table border="0" cellspacing="0" cellpadding="0" class="a">
	<tr>
		<td>* 커미션 : </td>
		<td><Font color="#3333FF"><%= CDbl(ijungsan.FCommission)*100 %> %</font></td>
	</tr>
	<tr>
		<td>* 총 건수 : </td>
		<td><Font color="#3333FF"><%= FormatNumber(ijungsan.FTotalCount,0) %></font></td>
	</tr>
	<tr>
		<td>* 정산대상 금액 : </td>
		<td ><% = FormatNumber(ijungsan.FTotalJungsan,0)  %></td>
	</tr>
	<tr>
		<td>* 정산예정 금액 : </td>
		<td ><% = FormatNumber((ijungsan.FTotalJungsan * CDbl(ijungsan.FCommission)),0)  %></td>
	</tr>
	</table>
	</td>
</tr>
<tr bgcolor="#DDDDDD">
	<td colspan="11" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="center">No</td>
	<td align="center">정산기간</td>
	<td align="center">건수</td>
	<td align="center">총금액</td>
	<td align="center">배송비</td>
	<td align="center">정산대상금액</td>
	<td align="center">정산금액</td>
	<td align="center">입금상태</td>
	<td align="center">세금일</td>
	<td align="center">입금일</td>
	<td align="center">내역받기</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr  bgcolor="#FFFFFF">
	<td colspan="11" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" >
	<input type="hidden" name="id" value="<%= ijungsan.FJungSanList(ix).Fmasterid %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="segumil" value="">
	<tr class="a"  bgcolor="#FFFFFF" >
		<td align="center"><% =ix + 1  %></td>
		<td align="center"><a href="oldjungsanlist_detail_view.asp?masterid=<% =ijungsan.FJungSanList(ix).Fmasterid %>&extsitename=<% =extsitename  %>"><%= ijungsan.FJungSanList(ix).FTotaldate %> ~ <%= ijungsan.FJungSanList(ix).FTotaldate2 %></a></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FTotalno %></td>
		<td align="right"><%= Formatnumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaldeasang,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaljungsansum,0) %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="center">
		<% if ijungsan.FJungSanList(ix).FIpkumDiv = 0 then %>
		  <font color="blue">No</font>
		 <% else %>
          <font color="red">Yes</font>
		 <% end if %>
		</td>
		<td align="center">
		<% if IsNULL(ijungsan.FJungSanList(ix).FSegumil) then %>
		<img src="/images/calicon.gif" border="0" id="sgDt<%= ijungsan.FJungSanList(ix).Fmasterid %>_trigger" style="cursor:pointer;" />
		<script>
			new Calendar({
				inputField : "sgDt", trigger    : "sgDt<%= ijungsan.FJungSanList(ix).Fmasterid %>_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					date = Calendar.printDate(date, "%Y-%m-%d");
					var frm = document.frmBuyPrc_<%= ix %>;
					if (confirm('세금일 : ' + date + ' OK?')){
						frm.id.value = "<%= ijungsan.FJungSanList(ix).Fmasterid %>";
						frm.mode.value = "segumil";
						frm.segumil.value = date;
						frm.submit();
					}
					this.hide();
				}, bottomBar: true
			});
		</script>
		<% else %>
		<%= ijungsan.FJungSanList(ix).FSegumil %>
		<% end if %>
		</td>
		<td align="center">
		 <% if ijungsan.FJungSanList(ix).FIpkumDiv = 0 then %>
		   <input type="button" value="정산하기" onclick="javascript:jungsanopen(<%= ijungsan.FJungSanList(ix).Fmasterid %>);">
		 <% else %>
           <% =FormatDateTime(ijungsan.FJungSanList(ix).FIpkumDate,2)  %>
		 <% end if %>
		</td>
		<td align="center">
		 	<a href="oldjungsanlist_detail_excel.asp?masterid=<% =ijungsan.FJungSanList(ix).Fmasterid %>&extsitename=<% =extsitename  %>"><img src="/images/icon_excel.gif" /></a>
		</td>
	</tr>
	</form>
	<% next %>
	<tr  bgcolor="#FFFFFF">
		<td colspan="13" height="30" align="center">
		<% if ijungsan.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ijungsan.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ijungsan.StarScrollPage to ijungsan.FScrollCount + ijungsan.StarScrollPage - 1 %>
			<% if ix>ijungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ijungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set ijungsan = nothing
%>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" crossorigin="anonymous" referrerpolicy="no-referrer" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<style type="text/css">
	.select2-container .select2-selection--single {height:17px;}
	.select2-container--default .select2-selection--single .select2-selection__rendered {line-height:16px;}
	.select2-container--default .select2-selection--single .select2-selection__arrow {height: 15px;}
</style>
<script>
$(function() {
	$("select[name=extsitename]").select2();
});
</script>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->