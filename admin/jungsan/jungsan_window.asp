<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim page
dim ijungsan
dim masterid
dim jungsanupdate

jungsanupdate = request("jungsanupdate")

nowdate = Left(CStr(now()),10)

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

masterid = request("masterid")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

set ijungsan = new CUpcheJungSan


if jungsanupdate = "update" then

ijungsan.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ijungsan.Fmasterid = masterid
ijungsan.PartnerOldJungSanDeasangUpdate


response.write "<script language='JavaScript'>self.close();</script>"
end if

%>

<html>
<head>
<title> 입금일 </title>
</head>

<body>
<form method=post name="form">
<input type="hidden" name="jungsanupdate" value="update">
<table width="250" height="150" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="center">정산하기</td>
</tr>
<tr>
	<td align="center">날짜 :
		<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr>
	<td align="center"><input type="submit" value="정산"></td>
</tr>
</table>
</form>
</body>
</html>
<%
set ijungsan = nothing
%>

<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->