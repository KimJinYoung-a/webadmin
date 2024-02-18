<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->

<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, Stype
dim nowdate,searchnextdate
dim page
dim extsitename
dim premonth, recomission

nowdate = Left(CStr(now()),10)
recomission = request("recomission")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

extsitename = request("extsitename")
Stype = request("Stype")
page = request("page")
if (page="") then page=1
if (Stype="") then Stype="B"

if (yyyy1="") then
	premonth = DateAdd("d",Cdate(Left(nowdate,4) + "-" + Mid(nowdate,6,2) + "-" + "01"),-1)
	yyyy1 = Left(premonth,4)
	mm1   = Mid(premonth,6,2)
	dd1   = "01"

	yyyy2 = Left(premonth,4)
	mm2   = Mid(premonth,6,2)
	dd2   = Mid(premonth,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


dim ijungsan
set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=7000
ijungsan.getDefaultInfo extsitename

ijungsan.FrectSiteName = extsitename
ijungsan.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ijungsan.FRectRegEnd   = searchnextdate
ijungsan.FRectDateSearchType = Stype
ijungsan.JungSanDeasangList
'ijungsan.JungSanDeasangList_OLD

dim ix
dim bufsum, deasangsum, amountsum, deliversum
bufsum =0
deasangsum =0
amountsum =0
deliversum =0
%>
<script language='javascript'>
function AnCheckNJungsan(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 내역이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문으로 정산자료를 작성하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.userid.value = upfrm.userid.value + "|" + frm.userid.value;
					upfrm.buyname.value = upfrm.buyname.value + "|" + frm.buyname.value;
					upfrm.totalsum.value = upfrm.totalsum.value + "|" + frm.totalsum.value;
					upfrm.deasangsum.value = upfrm.deasangsum.value + "|" + frm.deasangsum.value;
					upfrm.beasongpay.value = upfrm.beasongpay.value + "|" + frm.beasongpay.value;
					upfrm.jungsansum.value = upfrm.jungsansum.value + "|" + frm.jungsansum.value;
				}
			}
		}
		upfrm.submit();
	}
}

function ViewOrderDetail(frm){
    frm.target = 'orderdetail';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="1000" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" width="600">
		사이트:
		<% drawSelectBoxPartner "extsitename",extsitename %>
		<br>
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="radio" name="Stype" value="B" <%= CHKIIF(Stype="B","checked","") %> > 배송일 기준
		<input type="radio" name="Stype" value="R" <%= CHKIIF(Stype="R","checked","") %> > 주문일 기준
		
		<br>
		<!--
		2005-10-10 ~ 2005-10-19 이벤트 기간중 수수료 <input type=text name=recomission value="<%= recomission %>" size=2 maxlength=3 >%적용
		-->
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="1000" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="13">
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
		<td ><DIV ID="deasangsum">0</DIV></td>
	</tr>
	<tr>
		<td>* 정산예정 금액 : </td>
		<td ><DIV ID="amountsum">0</DIV></td>
	</tr>
	<tr>
		<td>* 배송비 금액 : </td>
		<td ><DIV ID="deliversum">0</DIV></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
    <td colspan="5" align="left">현재건색건 : <%= ijungsan.FResultCount %> (최대 <%= ijungsan.FPageSize %> 건 검색가능)</td>
	<td colspan="8" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr>
	<td colspan="13">
		<input type="button" value="전체선택" onClick="AnSelectAllFrame(true)">
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="선택사항정산" onclick="AnCheckNJungsan()">
	</td>
</tr>
<tr >
	<td width="30" align="center">선택</td>
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">포장.배송료</td>
	<td width="90" align="center">정산대상금액</td>
	<td width="90" align="center">정산금액</td>
	<td width="100" align="center">주문일</td>
	<td width="100" align="center">배송일</td>
	<td width="100" align="center">업체주문번호</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ijungsan.FJungSanList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ijungsan.FJungSanList(ix).FOrderSerial %>">
	<input type="hidden" name="userid" value="<%= ijungsan.FJungSanList(ix).FUserID %>">
	<input type="hidden" name="buyname" value="<%= ijungsan.FJungSanList(ix).FBuyName %>">
	<input type="hidden" name="totalsum" value="<%= ijungsan.FJungSanList(ix).FSubTotalPrice %>">
	<input type="hidden" name="beasongpay" value="<%= ijungsan.FJungSanList(ix).FBeasongPay %>">
	<input type="hidden" name="deasangsum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay %>">
	<% if (recomission<>"") and (ijungsan.FJungSanList(ix).FRegDate>="2005-10-10") and (ijungsan.FJungSanList(ix).FRegDate<"2005-10-20") then %>
	<input type="hidden" name="jungsansum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay * CDbl(recomission/100) %>">
	<% else %>
	<input type="hidden" name="jungsansum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay * CDbl(ijungsan.FCommission) %>">
	<% end if %>
	<tr class="a">
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td align="center"><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ijungsan.FJungSanList(ix).FOrderSerial %>)" class="zzz"><%= ijungsan.FJungSanList(ix).FOrderSerial %></a></td>
		<% if ijungsan.FJungSanList(ix).FUserID<>"" then %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FUserID %></td>
		<% else %>
		<td align="center">&nbsp;</td>
		<% end if %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FBuyName %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FDeasangPay,0) %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
			deliversum = deliversum + ijungsan.FJungSanList(ix).FBeasongPay
		 %>
		<% if (recomission<>"") and (ijungsan.FJungSanList(ix).FRegDate>="2005-10-10") and (ijungsan.FJungSanList(ix).FRegDate<"2005-10-20") then %>
		<td align="right" bgcolor="#CC3333"><%= FormatNumber(bufsum* CDbl(recomission/100),0) %></td>
		<% else %>
		<td align="right"><%= FormatNumber(bufsum* CDbl(ijungsan.FCommission),0) %></td>
		<% end if %>
		<td align="center"><%= Left(ijungsan.FJungSanList(ix).FRegDate,10) %></td>
		<td align="center"><%= Left(ijungsan.FJungSanList(ix).Fbeadaldate,10) %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).Fauthcode & ijungsan.FJungSanList(ix).Fpaygatetid %></td>
	</tr>
	</form>
	<% next %>
	<tr>
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
<form name="frmArrupdate" method="post" action="jungsanmaker.asp">
<input type="hidden" name="extsitename" value="<%= extsitename %>">
<input type="hidden" name="commission" value="<%= ijungsan.FCommission %>">
<input type="hidden" name="startdate" value="<%= yyyy1 + "-" + Format00(2,mm1) + "-" + Format00(2,dd1) %>">
<input type="hidden" name="enddate" value="<%= yyyy2 + "-" + Format00(2,mm2) + "-" + Format00(2,dd2) %>">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="userid" value="">
<input type="hidden" name="buyname" value="">
<input type="hidden" name="totalsum" value="">
<input type="hidden" name="deasangsum" value="">
<input type="hidden" name="beasongpay" value="">
<input type="hidden" name="jungsansum" value="">
</form>
<%
set ijungsan = nothing
%>
<script language='javascript'>
	deasangsum.innerText = '<%= FormatNumber(deasangsum,0) %>';
	amountsum.innerText = '<%= FormatNumber(amountsum,0) %>';
	deliversum.innerText = '<%= FormatNumber(deliversum,0) %>';
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->