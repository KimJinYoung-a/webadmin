<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드별매출
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
const Maxlines = 10

dim i
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim ojumun
dim fromDate,toDate
dim ordertype
dim makerid
dim oldlist, vPurchasetype
dim channelDiv

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
ordertype = request("ordertype")
makerid = request("makerid")
oldlist = request("oldlist")
vPurchasetype = request("purchasetype")
channelDiv  = NullFillWith(request("channelDiv"),"")

if (ordertype="") then ordertype="totalprice"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

set ojumun = new CJumunMaster
ojumun.FRectFromDate = fromDate
ojumun.FRectToDate = toDate
ojumun.FRectOrderType = ordertype
ojumun.FRectDesignerID = makerid
ojumun.FRectOldJumun = oldlist
ojumun.FPurchasetype = vPurchasetype
ojumun.FRectChannelDiv = channelDiv
ojumun.SearchSellrePort


dim sellcnt, selltotal, buytotal
dim itemcount, ordercount
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function checkform(frm)
{
	if(frm.oldlist.checked)
	{
		var date1 = new Date(frm.yyyy1.value,frm.mm1.value,frm.dd1.value);
		var date2 = new Date(frm.yyyy2.value,frm.mm2.value,frm.dd2.value);
		
		var years  = date2.getFullYear() - date1.getFullYear();
		var months = date2.getMonth() - date1.getMonth();
		var days   = date2.getDate() - date1.getDate();

		var chkmonth = years * 12 + months + (days >= 0 ? 0 : -1);

		if(chkmonth < 0)
		{
			alert("검색 기간 뒤에 날짜가 잘못되었습니다.");
			return false;
		}
		else if(chkmonth > 1)
		{
			alert("6개월 이전 검색은\n시작월과 마지막월 차이를 1개월 내로만 지정하세요.\n\n예: 2010-01-01 ~ 2010-02-01");
			return false;
		}
	}
	else
	{
	
	}
	
	//frm.submit();
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="upchesellamount.asp" onSubmit="return checkform(this);">
      <input type="hidden" name="showtype" value="showtype">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		*브랜드 :
		<% drawSelectBoxDesigner "makerid",makerid %>&nbsp;
		*검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        
        <input type="radio" name="ordertype" value="ea" <% if ordertype="ea" then response.write "checked" %>>수량순
		<input type="radio" name="ordertype" value="totalprice" <% if ordertype="totalprice" then response.write "checked" %>>매출순
		<input type="radio" name="ordertype" value="totalgain" <% if ordertype="totalgain" then response.write "checked" %>>수익순
		&nbsp;&nbsp;
		*구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		&nbsp;&nbsp;
                	*채널구분 
                	<select name="channelDiv">
                	<option value="" <%=CHKIIF(channelDiv="","selected","") %> >전체</option>
                	<option value="web" <%=CHKIIF(channelDiv="web","selected","") %> >웹</option>
                	<option value="jaehu" <%=CHKIIF(channelDiv="jaehu","selected","") %> >제휴</option>
                	<option value="mjaehu" <%=CHKIIF(channelDiv="mjaehu","selected","") %> >모바일제휴</option>
                	<option value="mobile" <%=CHKIIF(channelDiv="mobile","selected","") %> >모바일</option>
                	<option value="ipjum" <%=CHKIIF(channelDiv="ipjum","selected","") %> >상품입점</option>
                	</select>
                	<a href="javascript:image_view('http://webadmin.10x10.co.kr/admin/maechul/statistic/ch_gubun_exp.jpg');">[설명]</a>
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	<tr>
		<td class="a" colspan="2">
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" class="a" >
<tr>    
<td>주문일 기준, 반품주문건 포함, 배송비 제외</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00" class="a" >
    <tr align="center">
      <% if ojumun.FRectDesignerID<>"" then %>
      <td class="a"><font color="#FFFFFF">주문일</font></td>
      <% else %>
      <td class="a" width="30"><font color="#FFFFFF">순번</font></td>
	  <td class="a"><font color="#FFFFFF">브랜드 ID</font></td>
	  <% end if %>
	  <td width="80"><font color="#FFFFFF">상품가지수</font></td>
      <td width="80"><font color="#FFFFFF">상품갯수합</font></td>
      <td width="80"><font color="#FFFFFF">주문건수합</font></td>
      <td width="120"><font color="#FFFFFF">매출액(원)</font></td>
      <td width="120"><font color="#FFFFFF">매입액(원)</font></td>
      <td width="120"><font color="#FFFFFF">수익(원)</font></td>
      <td width="80"><font color="#FFFFFF">수익율(%)</font></td>
      <td width="80"><font color="#FFFFFF">주문객단가</font></td>
    </tr>
	<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center"  class="a">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
    <% for i=0 to ojumun.FResultCount - 1 %>
    <%
    sellcnt     = sellcnt + ojumun.FMasterItemList(i).Fsellcnt
    selltotal   = selltotal + ojumun.FMasterItemList(i).Fselltotal
    buytotal    = buytotal + ojumun.FMasterItemList(i).Fbuytotal
    itemcount   = itemcount + ojumun.FMasterItemList(i).Fitemcount
    ordercount  = ordercount + ojumun.FMasterItemList(i).Fordercount
    %>
    <tr bgcolor="#FFFFFF">
      <% if ojumun.FRectDesignerID<>"" then %>
      <td ><%= ojumun.FMasterItemList(i).FDate %></td>
      <% else %>
      <td ><%= i+1 %></td>
	  <td ><%= ojumun.FMasterItemList(i).Fmakerid %></td>
      <% end if %>
      
      <td align="center"><%= FormatNumber(ojumun.FMasterItemList(i).Fitemcount,0) %></td>
      <td align="center"><%= ojumun.FMasterItemList(i).Fsellcnt %></td>
	  <td align="center"><%= FormatNumber(ojumun.FMasterItemList(i).Fordercount,0) %></td>
	  <td align="right"><%= FormatNumber(ojumun.FMasterItemList(i).Fselltotal,0) %></td>
	  <td align="right"><%= FormatNumber(ojumun.FMasterItemList(i).Fbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(ojumun.FMasterItemList(i).Fselltotal-ojumun.FMasterItemList(i).Fbuytotal,0) %></td>
	  <td align="center">
		  <% if ojumun.FMasterItemList(i).Fselltotal<>0 then %>
		  <%= 100 - CLng(ojumun.FMasterItemList(i).Fbuytotal/ojumun.FMasterItemList(i).Fselltotal*100) %> 
		  <% end if %>
	  </td>
	  <td align="right">
	    <% if ojumun.FMasterItemList(i).Fordercount<>0 then %>
	        <%= FormatNumber(CLng(ojumun.FMasterItemList(i).Fselltotal/ojumun.FMasterItemList(i).Fordercount),0) %>
	    <% end if %>
	  </td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <% if ojumun.FRectDesignerID<>"" then %>
	  <td >Total</td>
	  <% else %>
	  <td colspan="2">Total</td>
	  <% end if %>
	  <td align="center"><%= FormatNumber(itemcount,0) %></td>
	  <td align="center"><%= FormatNumber(sellcnt,0) %></td>
	  <td align="center"><%= FormatNumber(ordercount,0) %></td>
	  <td align="right"><%= FormatNumber(selltotal,0) %></td>
	  <td align="right"><%= FormatNumber(buytotal,0) %></td>
	  <td align="right"><%= FormatNumber(selltotal-buytotal,0) %></td>
	  <td align="center">
	      <% if selltotal<>0 then %>
		  <%= 100 - CLng(buytotal/selltotal*100) %> 
		  <% end if %>
	  </td>
	  <td align="right">
	    <% if ordercount<>0 then %>
	        <%= FormatNumber(CLng(selltotal/ordercount),0) %>
	    <% end if %>
	  </td>
	</tr>
	<% end if %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
</body>
</html>
