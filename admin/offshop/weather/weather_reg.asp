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
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/weather/weather_cls.asp"-->

<%
Dim vIdx, cWeather, vWDate, vShopID, vShopName, vWeather, vComment, shopid, menupos
	menupos = requestCheckVar(Request("menupos"),10)
	vIdx = requestCheckVar(Request("idx"),10)
	
If vIdx <> "" Then
	Set cWeather = new COffShopWeather
	cWeather.FRectIdx = vIdx
	cWeather.GetOffShopWeatherView
	
	vWDate		= cWeather.FWDate
	vShopID		= cWeather.FShopID
	vShopName	= cWeather.FShopName
	vWeather	= cWeather.FWeather
	vComment	= cWeather.FComment
	Set cWeather = Nothing
End If

'직영/가맹점
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		vShopID = C_STREETSHOPID
	end if	
end if
	
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
function checkform(f)
{
	if(f.wdate.value == "")
	{
		alert("날짜를 선택하세요.");
		return false;
	}
	if(f.shopid.value == "")
	{
		alert("SHOP을 선택하세요.");
		return false;
	}
	if(!f.weather[0].checked && !f.weather[1].checked && !f.weather[2].checked && !f.weather[3].checked && !f.weather[4].checked && !f.weather[5].checked)
	{
		alert("날씨를 선택하세요.");
		return false;
	}
}
</script>
</head>
<body bgcolor="#F4F4F4">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<form name="frm" action="weather_process.asp" method="post" style="margin:0px;" onSubmit="return checkform(this);">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td bgcolor="#E6E6E6" align="center" width="70">날 짜</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="wdate" size="10" maxlength=10 readonly value="<%=vWDate%>">
		<a href="javascript:calendarOpen(frm.wdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center" width="70">SHOP</td>
	<td bgcolor="#FFFFFF">
	<% If vIdx <> "" Then %>
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(vShopID) <> "1" and vShopID <> "" then %>
				<%=vShopID%><input type="hidden" name="shopid" value="<%= vShopID %>">
			<% else %>
				<% 'drawSelectBoxOffShop "shopid",vShopID %>
				<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",vShopID, "21") %>
			<% end if %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",vShopID %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",vShopID, "21") %>
		<% end if %>
	<% Else %>
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(vShopID) <> "1" and vShopID <> "" then %>
				<%=vShopID%><input type="hidden" name="shopid" value="<%= vShopID %>">
			<% else %>
				<b><font color="blue">※ 다중 선택시 Ctrl 키를 누른채 마우스로 클릭하시면 됩니다.</font></b>
				<% drawSelectBoxMultiOffShop "shopid",vShopID %>
			<% end if %>
		<% else %>
			<b><font color="blue">※ 다중 선택시 Ctrl 키를 누른채 마우스로 클릭하시면 됩니다.</font></b>
			<% drawSelectBoxMultiOffShop "shopid",vShopID %>
		<% end if %>
	<% End If %>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center" width="70">날 씨</td>
	<td bgcolor="#FFFFFF" align="center">
		<table cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" style="padding-right:15px;"><img src="/images/weather/1.gif" width="50" style="cursor:pointer;" onClick="frm.weather[0].checked=true;"></td>
			<td align="center" style="padding-right:15px;"><img src="/images/weather/2.gif" width="50" style="cursor:pointer;" onClick="frm.weather[1].checked=true;"></td>
			<td align="center" style="padding-right:15px;"><img src="/images/weather/3.gif" width="50" style="cursor:pointer;" onClick="frm.weather[2].checked=true;"></td>
			<td align="center" style="padding-right:15px;"><img src="/images/weather/4.gif" width="50" style="cursor:pointer;" onClick="frm.weather[3].checked=true;"></td>
			<td align="center" style="padding-right:15px;"><img src="/images/weather/5.gif" width="50" style="cursor:pointer;" onClick="frm.weather[4].checked=true;"></td>
			<td align="center"><img src="/images/weather/6.gif" width="50" style="cursor:pointer;" onClick="frm.weather[5].checked=true;"></td>
		</tr>
		<tr>
			<td align="center" style="padding-right:15px;"><input type="radio" name="weather" value="1" <%=CHKIIF(vWeather="1","checked","")%>></td>
			<td align="center" style="padding-right:15px;"><input type="radio" name="weather" value="2" <%=CHKIIF(vWeather="2","checked","")%>></td>
			<td align="center" style="padding-right:15px;"><input type="radio" name="weather" value="3" <%=CHKIIF(vWeather="3","checked","")%>></td>
			<td align="center" style="padding-right:15px;"><input type="radio" name="weather" value="4" <%=CHKIIF(vWeather="4","checked","")%>></td>
			<td align="center" style="padding-right:15px;"><input type="radio" name="weather" value="5" <%=CHKIIF(vWeather="5","checked","")%>></td>
			<td align="center"><input type="radio" name="weather" value="6" <%=CHKIIF(vWeather="6","checked","")%>></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center" width="70">코맨트</td>
	<td bgcolor="#FFFFFF"><textarea name="comment" cols="50" rows="5"><%=vComment%></textarea></td>
</tr>
</table>
<br><input type="submit" value="저 장" class="button">
</form>

</body>
</html>

<%
Sub drawSelectBoxMultiOffShop(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>" multiple size="20">
     <%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end sub
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->