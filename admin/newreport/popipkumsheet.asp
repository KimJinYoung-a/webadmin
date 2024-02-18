<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim yyyy1,mm1,dd1
dim yyyy2,mm2,dd2
dim ipkumstate,tenbank,ipkumname,page

ipkumstate=request("ipkumstate")
tenbank=request("tenbank")
ipkumname=request("ipkumname")
page=request("page")
if page="" then page=1

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))-1
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if 

dim ipkum,i,ix
set ipkum = new IpkumChecklist

ipkum.FCurrpage=page
ipkum.cksheep=1
ipkum.FScrollCount = 5
ipkum.ipkumstate=ipkumstate
ipkum.Ctenbank=tenbank
ipkum.ipkumname=ipkumname

ipkum.yyyy1=yyyy1
ipkum.mm1=mm1
ipkum.dd1=dd1
ipkum.yyyy2=yyyy2
ipkum.mm2=mm2
ipkum.dd2=dd2

ipkum.Getipkumlist

response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + yyyy1 + mm1 + dd1 + yyyy2 + mm2 + dd2 + tenbank + ipkumname + ".xls"
%>
<html xmlns:x="urn:schemas-microsoft-com:office:excel">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<style>
  .big_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:18.0pt;
    font-weight:700;
    }
  .mid_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:12.0pt;
    font-weight:700;
    }
  .title_center
	{mso-style-parent:style0;
	text-align:center;
	vertical-align:middle;
	white-space:normal;}
  .normal
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	}
  .normal_b
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	font-weight:700;
	}
  .currency
	{mso-style-parent:style0;
 	mso-number-format:"\#\,\#\#0\.00_\)\;\[Red\]\\\(\#\,\#\#0\.00\\\)";
	border:0.5pt solid black;
	white-space:normal;}
   .Format_Y1
	{mso-style-parent:style0;
	mso-number-format:"yyyy\0022\/\0022m\0022\/\0022d\;\@";
 	white-space:normal;}
   .Format_Y2
	{mso-style-parent:style0;
	mso-number-format:"yyyy\/mm\;\@";
	text-align:center;
	border:0.5pt solid black;
 	white-space:normal;}
   .Format_number
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:8.0pt;
	}
   .Format_number_L
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:12.0pt;
	}
  .Format_T1
	{mso-style-parent:style0;
	mso-number-format:"hh\:mm\:ss\;\@";
	text-align:center;
 	white-space:normal;}  </style>
</head>
<body leftmargin="10">
<table width=700 cellspacing=0 cellpadding=1 border=0>
<tr bgcolor="#FFFFFF">
 	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>Idx</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>은행</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>날짜</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>구분</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>입금자</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>출금액</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>입금액</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>잔액</td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>적요</td>
 </tr>
 <% if ipkum.FResultCount<1 then %>
<% else %>
<% for i=0 to ipkum.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fidx %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Ftenbank %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= left(ipkum.Fipkumitem(i).FBankdate,10) %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fgubun %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fipkumuser %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fchulkumsum %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fipkumsum %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center><%= ipkum.Fipkumitem(i).Fremainsum %></td>
	<td height=20 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;' align=center>&nbsp;</td>
</tr>
<% next %>
<% end if %>
</table>
</body>
</html>
<Script Language="Javascript">
self.close();
</Script>


<% set ipkum=nothing %> 

<!-- #include virtual="/lib/db/dbclose.asp" -->
