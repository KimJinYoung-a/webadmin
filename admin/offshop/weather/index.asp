<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'==========================================================================
'	Description: 매장날씨 관리
'	History: 2012.06.04 강준구 생성
'			 2012.06.12 한용민 수정
'==========================================================================
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/weather/weather_cls.asp"-->

<%
Dim shopid, cWeather, page, i, vWSDate, vWEDate
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	vWSDate = requestCheckVar(request("wstart"),10)
	vWEDate = requestCheckVar(request("wend"),10)

if page = "" then page = 1

'직영/가맹점
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if

Set cWeather = new COffShopWeather
cWeather.FPageSize = 20
cWeather.FRectShopID = shopid
cWeather.FCurrPage = page
cWeather.FRectWSDate = vWSDate
cWeather.FRectWEDate = vWEDate
cWeather.GetOffShopWeatherList
%>

<script language="javascript">

function weatherreg(idx){
	var iheight;
	if(idx == "")
	{
		iheight = 6;
	}
	else
	{
		iheight = 3;
	}
	
	var weatherreg = window.open('/admin/offshop/weather/weather_reg.asp?menupos=<%=menupos%>&idx='+idx,'weatherreg','width=500,height='+iheight+'00');
	weatherreg.focus();	
}

function commentview(idx){
	if(document.getElementById("comment"+idx+"").style.display == "block")
	{
		document.getElementById("comment"+idx+"").style.display = "none";
	}
	else
	{
		document.getElementById("comment"+idx+"").style.display = "block";
	}
}

function frmsubmit(page){
	frm.page.value = page;
	frm.submit();
}

</script>

<table cellpadding="3" cellspacing="1" class="a" width="100%" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>		
		&nbsp;&nbsp;&nbsp;
		날짜 : 
		<input type="text" name="wstart" size="10" maxlength=10 readonly value="<%=vWSDate%>">
		<a href="javascript:calendarOpen(frm.wstart);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="wend" size="10" maxlength=10 readonly value="<%=vWEDate%>">
		<a href="javascript:calendarOpen(frm.wend);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table cellpadding="3" cellspacing="1" class="a" width="100%" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				검색결과 : <b><%= cWeather.FTotalCount %></b>
				&nbsp;
				페이지 : <b><%= page %>/ <%= cWeather.FTotalPage %></b>
			</td>
			<td align="right"><input type="button" class="button" value="날씨등록" onClick="weatherreg('')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="150">날짜</td>
	<td width="200">SHOP</td>
	<td width="100">날씨</td>
	<td>코멘트</td>
	<td width="170"></td>
</tr>
<%
if cWeather.FresultCount>0 then
	
for i=0 to cWeather.FresultCount-1
%>
	<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		<td align="center"><%= cWeather.FItemList(i).FWDate %></td>
		<td align="center"><%= cWeather.FItemList(i).FShopName %>(<%= cWeather.FItemList(i).FShopID %>)</td>
		<td align="center"><img src="<%= cWeather.FItemList(i).FWeather %>" width="30"></td> 
		<td><% If cWeather.FItemList(i).FComment <> "" Then %><%= Replace(cWeather.FItemList(i).FComment,vbCrLf,"<br>") %><% End If %></td>  
		<td align="center">
			<input type="button" class="button" value="수 정" onClick="weatherreg('<%= cWeather.FItemList(i).FIdx %>');"> 
		</td>
	</tr>
	
<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cWeather.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= cWeather.StartScrollPage-1 %>');">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cWeather.StartScrollPage to cWeather.StartScrollPage + cWeather.FScrollCount - 1 %>
				<% if (i > cWeather.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cWeather.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cWeather.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% Set cWeather = Nothing %>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->