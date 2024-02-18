<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<!-- #include virtual="/lib/classes/tingordercls.asp"-->
<%
dim orderserial
dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim ckonlyOK

nowdate = Left(CStr(now()),10)

searchtype = request("searchtype")
searchrect = request("searchrect")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
ckonlyOK = request("ckonlyOK")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim cknodate,ckdelsearch,ckipkumdiv4
cknodate = request("cknodate")
ckdelsearch = request("ckdelsearch")
ckipkumdiv4 = request("ckipkumdiv4")
orderserial = request("orderserial")


dim page
dim ojumun, eventname

page = request("page")
if (page="") then page=1

eventname = request("eventname")

set ojumun = new CTingOrderMaster

ojumun.FPageSize = 50
ojumun.FCurrPage = page

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if ckdelsearch<>"on" then
	ojumun.FRectDelNoSearch="on"
end if

if searchtype="01" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="02" then
	ojumun.FRectReqName = searchrect
elseif searchtype="03" then
	ojumun.FRectUserID = searchrect
elseif searchtype="04" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="07" then
	ojumun.FRectItemID = searchrect
'elseif searchtype="06" then
'	ojumun.FRectSubTotalPrice = searchrect
end if

ojumun.FRectCheckOnlyOK = ckonlyOK

ojumun.FrectEventName = eventname
ojumun.SearchJumunList

dim ix

%>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value = ipage;
	document.frm.submit();
}
</script>

<table width="820" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		주문번호 : 
		<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
		&nbsp;
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		(<input type="checkbox" name="cknodate" <% if cknodate="on" then response.write "checked" %> >기간상관없음)
		<input type="checkbox" name="ckdelsearch" <% if ckdelsearch="on" then response.write "checked" %> >취소,삭제검색)
		<br>
		상품구분 :
		<select name="eventname">
		<option value="">선택</option>
		<option value="-" <% if eventname="-" then response.write "selected" %> >일반(배송)상품</option>
		<option value="maxmovie" <% if eventname="maxmovie" then response.write "selected" %> >맥스무비</option>
		<option value="telcoin" <% if eventname="telcoin" then response.write "selected" %> >텔코인</option>
		</select>
		검색조건 : 
		<select name="searchtype">
		<option value="">선택</option>
		<option value="01" <% if searchtype="01" then response.write "selected" %> >구매자</option>
		<option value="02" <% if searchtype="02" then response.write "selected" %> >수령인</option>
		<option value="03" <% if searchtype="03" then response.write "selected" %> >아이디</option>
		<option value="04" <% if searchtype="04" then response.write "selected" %> >입금자</option>
		<option value="07" <% if searchtype="07" then response.write "selected" %> >아이템ID</option>
		</select>
		<input type="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
		&nbsp;
		<input type="checkbox" name="ckonlyOK" <% if ckonlyOK="on" then response.write "checked" %> >정상건수만검색
		</td>
		<td class="a" align="right">			
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="820" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="10">
		총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp;총 팅Q : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotal,0) %></font>
		&nbsp;평균 팅Q : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotal,0) %></font>
	</td>
</tr>
<tr>
	<td colspan="10" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr >
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">UserID</td>
	<td width="80" align="center">구매자</td>
	<td width="80" align="center">수령인</td>
	<td width="80" align="center">팅Q</td>
	<td width="60" align="center">삭제여부</td>
	<td width="140" align="center">주문일</td>
	<td width="70">구매자HP</td>
	<td width="140">결과코드</td>
	<td width="140">수령지주소</td>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr>
	<td colspan="10" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<tr>
		<% if IsNull(ojumun.FMasterItemList(ix).FOrderSerial) then %>
		<td align="center">-</td>
		<% else %>
		<td align="center"><%= CStr(ojumun.FMasterItemList(ix).FOrderSerial) %></td>
		<% end if %>
		<td align="center"><%= ojumun.FMasterItemList(ix).FUserID %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTingQ,0) %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FCancelYn %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).Forderdate %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FResultCode %></td>
		<td align="left"><%= ojumun.FMasterItemList(ix).Freqaddr1 + " " + ojumun.FMasterItemList(ix).Freqaddr2 %></td>
	</tr>
	<% next %>
	<tr>
		<td colspan="13" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StarScrollPage-1 %>')">[pre]</a>
		<% else %> 
			[pre]
		<% end if %>
		
		<% for ix=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>
		
		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %> 
			[next]
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