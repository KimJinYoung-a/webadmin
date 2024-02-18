<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/DataMartItemsalecls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2,Param
dim fromDate,toDate,cdL,cdM,cdS,research,cd4
dim rectoldjumun,dategubun, ckMinus
dim catebase
dim ck2ndcate
dim topn

'dim ck_joinmall,ck_ipjummall,ck_pointmall

yyyy1   = RequestCheckVar(request("yyyy1"),4)
mm1     = RequestCheckVar(request("mm1"),2)
dd1     = RequestCheckVar(request("dd1"),2)
yyyy2   = RequestCheckVar(request("yyyy2"),4)
mm2     = RequestCheckVar(request("mm2"),2)
dd2     = RequestCheckVar(request("dd2"),2)
topn    = RequestCheckVar(request("topn"),10)
rectoldjumun    = RequestCheckVar(request("rectoldjumun"),10)
dategubun       = ""
research        = RequestCheckVar(request("research"),10)
ckMinus         = RequestCheckVar(request("ckMinus"),10)
catebase  = RequestCheckVar(request("catebase"),10)
''ck_joinmall     = RequestCheckVar(request("ck_joinmall"),10)
''ck_ipjummall    = RequestCheckVar(request("ck_ipjummall"),10)
''ck_pointmall    = RequestCheckVar(request("ck_pointmall"),10)

cdL = RequestCheckVar(request("cd1"),10)
cdM = RequestCheckVar(request("cd2"),10)
cdS = RequestCheckVar(request("cd3"),10)
cd4 = RequestCheckVar(request("cd4"),10)

ck2ndcate = RequestCheckVar(request("ck2ndcate"),10)

if research<>"on" then
    ''if ckMinus="" then ckMinus="1"
	'if ck_joinmall="" then ck_joinmall="on"
	'if ck_ipjummall="" then ck_ipjummall="on"
	'if dategubun="" then dategubun="D"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2,Month(now()))
if (dd1="") then dd1 = Format00(2,day(now()))

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Format00(2,Month(now()))
if (dd2="") then dd2 = Format00(2,day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

if (catebase="") then catebase="V"

Param = "&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&dategubun="&dategubun&"&ckMinus="&ckMinus&"&catebase="&catebase

dim topMax : topMax=3000
if (topn="") then topn=topMax
if not isNumeric(topn) then topn=topMax
if (topn>topMax) then topn=topMax

dim oReport
set oReport = new CDatamartItemSale
oReport.FPageSize=topn
oReport.FRectStartDate = fromDate
oReport.FRectEndDate = toDate
oReport.FRectDateGubun = dategubun
oReport.FRectIncludeMinus = ckMinus
oReport.FRectCD1 = cdL
oReport.FRectCD2 = cdM
oReport.FRectCD3 = cdS
oReport.FRectCD4 = cd4
oReport.FRectOldJumun = rectoldjumun
oReport.FRectByBestSell ="on"
oReport.FRectInclude2ndCate = ck2ndcate

if (catebase="V") then
    oReport.SearchMallSellrePortChannelByCurrentDispCateBest    ''신 전시카테고리 기준
elseif (catebase="C") then
    oReport.SearchMallSellrePortChannelByCurrentCateBest   ''현재(관리) 카테고리 기준
else
    oReport.SearchMallSellrePortChannelBest           '' 판매시(관리) 카테고리 기준
end if

'oreport.SearchMallSellrePortMonthlyChannel
dim i,p1,p2
dim prename, nextname
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,ch5,ch6,ch7,ch8,ch9,ch10,ch11


dim sellcnt, selltotal, buytotal
dim TTLsellcnt, TTLselltotal, TTLbuytotal
dim TTLorgTotal,TTLitemcostCouponNotApplied
%>
<script language='javascript'>
function subPage(cd1,cd2,cd3){
    frm.cd1.value=cd1;
    frm.cd2.value=cd2;
    frm.cd3.value=cd3;
    frm.submit();
}


function gridQuery(page){
    ///lib/util/gridResponse.asp?cmd=channelSellsum&stdt=2009-01-01&eddt=2009-06-16
    var Grid1 = document.all.TnGrid;
    var frm   = document.frm;
    var pagesize = 20;

    Grid1.setQueryUrl("<%= manageUrl %>/lib/util/gridResponse.asp");
    Grid1.clearParams();
	Grid1.addParam("cmd", "channelSellsum");
	Grid1.addParam("page", page);
	Grid1.addParam("pagesize", pagesize);

    Grid1.addParam("stdt", frm.yyyy1.value+"-"+frm.mm1.value+"-"+frm.dd1.value);
    Grid1.addParam("eddt", frm.yyyy2.value+"-"+frm.mm2.value+"-"+frm.dd2.value);

	Grid1.addParam("ckMinus",frm.ckMinus.value);

	//ckMinus
	//Grid1.getWebData(parsingType);
	try {
	    Grid1.getWebData(1);
	}catch(e){
	    alert(e.description);
	}

}

function chkComp(comp){
    if (comp.value=="C"){
        comp.form.ck2ndcate.disabled=false;
    }else{

        comp.form.ck2ndcate.checked=false;
        comp.form.ck2ndcate.disabled=true;
    }
}
</script>
<!-- script language="JavaScript" src="/js/tnGrid.js"></script -->

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >

		주문일  :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<select name="ckMinus">
		<option value="" >반품포함
		<option value="1" <%= CHKIIF(ckMinus="1","selected","") %> >반품제외
		<option value="2" <%= CHKIIF(ckMinus="2","selected","") %> >반품주문만
		</select>

		&nbsp;
		카테고리 매출 기준:
		<input type="radio" name="catebase" value="V" <%= CHKIIF(catebase="V","checked","") %> onClick="chkComp(this)">현재(<strong>전시</strong>)카테고리
		<input type="radio" name="catebase" value="S" <%= CHKIIF(catebase="S","checked","") %> onClick="chkComp(this)">판매시(관리)카테고리
		<input type="radio" name="catebase" value="C" <%= CHKIIF(catebase="C","checked","") %> onClick="chkComp(this)">현재(관리)카테고리

		<select name="ck2ndcate">
		<option value="">기본카테고리
		<option value="All" <%= CHKIIF(ck2ndcate="All","selected","") %>>기본+추가카테고리
		<option value="OnlyA" <%= CHKIIF(ck2ndcate="OnlyA","selected","") %>>추가카테고리
		</select>
		<!--
		<input type="checkbox" name="ck2ndcate" <%= CHKIIF(ck2ndcate="on","checked","") %> <%= CHKIIF(catebase="S","disabled='true'","") %>>추가 카테고리 포함
		-->
		<td class="a" align="right" rowspan="2">
			<a href="javascript:frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	<tr>
	    <td class="a" >
	    검색갯수 <input type="text" name="topn" value="<%=topn%>" size="4" maxlength="4">

		<% if (catebase="V") then %>
	    1뎁스 <input type="text" name="cd1" value="<%=cdL%>" size="4" maxlength="4">
	    2뎁스 <input type="text" name="cd2" value="<%=cdM%>" size="4" maxlength="4">
	    3뎁스 <input type="text" name="cd3" value="<%=cdS%>" size="4" maxlength="4">
	    4뎁스 <input type="text" name="cd4" value="<%=cd4%>" size="4" maxlength="4">
		<% else %>
	    cdL <input type="text" name="cd1" value="<%=cdL%>" size="4" maxlength="4">
	    cdM <input type="text" name="cd2" value="<%=cdM%>" size="4" maxlength="4">
	    cdS <input type="text" name="cd3" value="<%=cdS%>" size="4" maxlength="4">
		<% end if %>

	    &nbsp;(최대 <%=topMax%> 건 검색됨)
	    </td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" class="a" >
<tr>
	<td>* 보너스 쿠폰, 마일리지 제외됨. 1시간 지연 데이터. 배송비 매출은 제외됨.</td>
</tr>
</table>

<!--
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" class="a" >
<tr align="center">
    <td>
    <script language='javascript'>
    DrawTnGridTag("TnGrid","100%",400);
    </script>
    </td>
</tr>
</table>
-->


<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#EFBE00" class="a" >
    <tr align="center">
      <td width="160" class="a"><font color="#FFFFFF">브랜드</font></td>
      <td width="80"  ><font color="#FFFFFF">상품코드</font></td>
      <td ><font color="#FFFFFF">상품명</font></td>
      <td ><font color="#FFFFFF">옵션명</font></td>
      <td width="80" class="a"><font color="#FFFFFF">건수<br>(상품수)</font></td>
      <% if (NOT C_InspectorUser) then %>
      <td width="90" class="a"><font color="#FFFFFF">소비자가</font></td>
      <td width="90" class="a"><font color="#FFFFFF">할인금액</font></td>
      <td width="90" class="a"><font color="#FFFFFF">판매가<br>(할인가)</font></td>
      <td width="90" class="a"><font color="#FFFFFF">상품쿠폰<br>사용액</font></td>
      <% end if %>
      <td width="90" class="a"><font color="#FFFFFF"><strong>구매총액</strong><% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></font></td>
      <td width="90" class="a"><font color="#FFFFFF">매입액</font></td>
      <td width="90" class="a"><font color="#FFFFFF">수익</font></td>
      <td width="80" class="a"><font color="#FFFFFF">수익율</font></td>
    </tr>
	<% for i=0 to oreport.FResultCount-1 %>
	<%

		sellcnt		=	sellcnt + oreport.FItemList(i).Fsellcnt
		selltotal	=	selltotal + oreport.FItemList(i).Fselltotal
		buytotal	=	buytotal + oreport.FItemList(i).Fbuytotal
	%>
	<tr bgcolor="#FFFFFF">
	  <td align="left">
	    <%= oReport.FItemList(i).Fmakerid %>
	  </td>
	  <td ><%= oReport.FItemList(i).Fitemid %></td>
	  <td ><%= oReport.FItemList(i).Fitemname %></td>
	  <td ><%= oReport.FItemList(i).Foptionname %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum-oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum-oreport.FItemList(i).Fselltotal) %></td>
	  <% end if %>
	  <td align="right" bgcolor="#7CCE76"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal,0) %></td>
	  <td align="center">
	  <% if oreport.FItemList(i).Fselltotal<>0 then %>
	  	<%= 100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	</tr>


	<%
		TTLsellcnt	= TTLsellcnt + sellcnt
		TTLselltotal= TTLselltotal + selltotal
		TTLbuytotal = TTLbuytotal + buytotal
        TTLorgTotal = TTLorgTotal + oreport.FItemList(i).ForgitemcostSum
        TTLitemcostCouponNotApplied = TTLitemcostCouponNotApplied + oreport.FItemList(i).FitemcostCouponNotAppliedSum

		sellcnt = 0
		selltotal = 0
		buytotal = 0
	%>

	<% next %>
	<tr bgcolor="#FFFFFF"><td colspan="13"></td></tr>
	<tr bgcolor="#FFFFFF">
	  <td align="center">Total</td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td align="right"><%= FormatNumber(TTLsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal-TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied-TTLselltotal) %></td>
	  <% end if %>
	  <td align="right" bgcolor="#7CCE76"><%= FormatNumber(TTLselltotal,0) %></td>
	  <td align="right"><%= FormatNumber(TTLbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(TTLselltotal-TTLbuytotal,0) %></td>
	  <td align="center">
	  <% if TTLselltotal<>0 then %>
	  	<%= 100-CLng(TTLbuytotal/TTLselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	</tr>

</table>


<%
set oreport = Nothing
%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->