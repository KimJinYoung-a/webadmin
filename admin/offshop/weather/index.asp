<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'==========================================================================
'	Description: ���峯�� ����
'	History: 2012.06.04 ���ر� ����
'			 2012.06.12 �ѿ�� ����
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

'����/������
if (C_IS_SHOP) then
	
	'/���α��� ���� �̸�
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
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>		
		&nbsp;&nbsp;&nbsp;
		��¥ : 
		<input type="text" name="wstart" size="10" maxlength=10 readonly value="<%=vWSDate%>">
		<a href="javascript:calendarOpen(frm.wstart);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="wend" size="10" maxlength=10 readonly value="<%=vWEDate%>">
		<a href="javascript:calendarOpen(frm.wend);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
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
				�˻���� : <b><%= cWeather.FTotalCount %></b>
				&nbsp;
				������ : <b><%= page %>/ <%= cWeather.FTotalPage %></b>
			</td>
			<td align="right"><input type="button" class="button" value="�������" onClick="weatherreg('')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="150">��¥</td>
	<td width="200">SHOP</td>
	<td width="100">����</td>
	<td>�ڸ�Ʈ</td>
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
			<input type="button" class="button" value="�� ��" onClick="weatherreg('<%= cWeather.FItemList(i).FIdx %>');"> 
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
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<% Set cWeather = Nothing %>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->