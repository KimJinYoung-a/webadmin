<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "회원가입현황"

dim settle
settle= request("settle")

if settle="" then
	settle = "D"
end if

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	date1 = dateAdd("m",-1,nowdate)
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)
	
	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)
	
	Edate = Left(CStr(nowdate+1),10)
end if	
%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<tr>
		<td class="a">
		조회기간 : 
			<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" width="100"><a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	<tr>
		<td class="a">
		그래프 : 
			<input type="radio" name="settle" value="M" <% if settle="M" then response.write "checked" %> >Year 
            <input type="radio" name="settle" value="D" <% if settle="D" then response.write "checked" %> >Month
            <input type="radio" name="settle" value="W" <% if settle="W" then response.write "checked" %> >Week
            <input type="radio" name="settle" value="T" <% if settle="T" then response.write "checked" %> >Time
			
		</td>
		<td></td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
