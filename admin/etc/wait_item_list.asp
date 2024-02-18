<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<%
dim iting,itemid,itemname
dim page
dim yyyy1,mm1,nowdate

nowdate = Left(CStr(now()),10)

yyyy1 = request("yyyy1")
mm1 = request("mm1")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2) + 1
end if

itemid = request("itemid")
itemname = request("itemname")
page = request("page")


page = request("page")
if page="" then page = 1

set iting = new CTingWaitItemList
iting.FPageSize = 50
iting.FCurrPage = page
iting.FRectPropmon = yyyy1 & "-" & mm1
iting.WaitItemList

dim ix
%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		상품ID :
		<input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="9" class="input_b">
		상품명 :
		<input type="text" name="itemname" value="<%= itemname %>" size="12" maxlength="32" class="input_b">
		제안월 : <% DrawYMBox yyyy1,mm1 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="12" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(iting.FTotalPage,0) %> count: <%= FormatNumber(iting.FTotalCount,0) %></td>
</tr>
<tr>
	<form name="frmttl" onsubmit="return false;">
	<td colspan="12" height="30"><input type="button" value="전체선택" onClick="AnSelectAllFrame(true)" class="button">&nbsp;<input type="button" value="선택상품저장" onClick="AnItemviewsetSaveAll()" class="button"></td>
	</form>
</tr>
<tr>
	<td align="center">선택</td>
	<td align="center">제안월</td>
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">상품명</td>
	<td align="center">제안가격</td>
	<td align="center">팅포인트(정)</td>
	<td align="center">팅포인트(비)</td>
	<td align="center">구매구분</td>
	<td align="center">한정판매</td>
	<td align="center">총한정수량</td>
	<td align="center">선택결과</td>
</tr>
<tr>
	<td colspan="12" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
<% if iting.FresultCount < 1 then %>
<tr>
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% for ix = 0 to iting.FresultCount - 1 %>
<form name="frmBuyPrc_<%= iting.FTingList(ix).FItemID %>" method="post" onSubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= iting.FTingList(ix).FItemID %>">
<tr height="20">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= iting.FTingList(ix).Fpropmon %></td>
	<td align="center"><%= iting.FTingList(ix).FItemID %></td>
	<td align="center"><img src="<%= iting.FTingList(ix).FImageSmall %>" width="50" height="50" border=0></td>
	<td>&nbsp;<%= iting.FTingList(ix).FItemName %></td>
	<td align="center"><%= FormatNumber(iting.FTingList(ix).Fpropcost,0) %>원</td>
	<td align="center"><input type="text" name="tingpoint" value="<%= iting.FTingList(ix).FTingPoint %>" size=6 class="input_b"></td>
	<td align="center"><input type="text" name="tingpoint_b" value="<%= iting.FTingList(ix).FTingPoint_B %>" size=6 class="input_b"></td>
	<td align="center">
		<select name="userclass">
			<option value="A" <% if iting.FTingList(ix).FUserClass = "A" then response.write "selected" %> >정팅,준팅</option>
			<option value="Y" <% if iting.FTingList(ix).FUserClass = "Y" then response.write "selected" %> >정팅</option>
			<option value="N" <% if iting.FTingList(ix).FUserClass = "N" then response.write "selected" %> >정팅,준팅,노팅</option>
		</select>
	</td>
	<td align="center">
		<select name="limitdiv">
			<option value="0" <% if iting.FTingList(ix).FLimitDiv = "0" then response.write "selected" %> >비한정판매</option>
			<option value="1" <% if iting.FTingList(ix).FLimitDiv = "1" then response.write "selected" %> >수량한정</option>
			<option value="2" <% if iting.FTingList(ix).FLimitDiv = "2" then response.write "selected" %> >일별한정</option>
			<option value="3" <% if iting.FTingList(ix).FLimitDiv = "3" then response.write "selected" %> >월별한정</option>
		</select>
	</td>
	<td align="center"><input type="text" name="limitea" value="<%= iting.FTingList(ix).FLimitea %>" size=6 class="input_b"></td>
	<td align="center">
	<select name="selectitem">
	<option value="Y" <% if iting.FTingList(ix).Fselectitem = "Y" then response.write "selected" %> style="background-color:#CCFFFF;">선택</option>
	<option value="N" <% if iting.FTingList(ix).Fselectitem = "N" then response.write "selected" %>>미선택</option>
	</select>
	</td>
</tr>
<tr>
	<td colspan="12" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
</form>
<% next %>
<% end if %>
<tr>
	<td colspan="12" align="center">
	<% if iting.HasPreScroll then %>
		<a href="?page=<%= iting.StarScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[pre]</a>
	<% else %>
	<% end if %>

	<% for ix=0 + iting.StarScrollPage to iting.FScrollCount + iting.StarScrollPage - 1 %>
		<% if ix > iting.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(ix) then %>
		<font color="red">[<%= ix %>]</font>
		<% else %>
		<a href="?page=<%= ix %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[<%= ix %>]</a>
		<% end if %>
	<% next %>

	<% if iting.HasNextScroll then %>
		<a href="?page=<%= ix %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[next]</a>
	<% else %>
	<% end if %>
	</td>
</tr>

<tr>
	<td colspan="12" height="20">
</tr>
<form name="frmArrupdate" method="post" action="wait_item_update.asp">
<input type="hidden" name="itemidlist">
<input type="hidden" name="tingpointlist">
<input type="hidden" name="tingpoint_blist">
<input type="hidden" name="userclasslist">
<input type="hidden" name="limitdivlist">
<input type="hidden" name="limitealist">
<input type="hidden" name="selectitemlist">
</form>
</table>
<%
set iting = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->