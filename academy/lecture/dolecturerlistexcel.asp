<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
'==============================================================================
''corpse2 추가 강사리스트 엑셀 다운로드 2017-03-14
'==============================================================================
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim page, i
dim opartner
dim mduserid, rect, isusing, research

page        = requestCheckVar(request("page"),10)
mduserid    = requestCheckVar(request("mduserid"),32)
rect        = requestCheckVar(request("rect"),32)
isusing     = requestCheckVar(request("isusing"),10)
research    = requestCheckVar(request("research"),10)

if page="" then page=1
if isusing="" and research="" then isusing="on"

set opartner = new CPartnerUser
opartner.FCurrpage = page
opartner.FPageSize = 1000
opartner.FRectIsUsing = isusing
opartner.FRectInitial=rect
opartner.GetAcademyPartnerList

Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=LectureList_" & Left(CStr(now()),10) & ".xls"
Response.CacheControl = "public"

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>강사리스트</title>
<style>
<!--
table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:.75in .7in .75in .7in;
	mso-header-margin:.3in;
	mso-footer-margin:.3in;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-char-type:none;
	display:none;}
tr
	{mso-height-source:auto;
	mso-ruby-visibility:none;}
col
	{mso-width-source:auto;
	mso-ruby-visibility:none;}
br
	{mso-data-placement:same-cell;}
ruby
	{ruby-align:left;}
.style0
	{mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	border:none;
	mso-protection:locked visible;
	mso-style-name:표준;
	mso-style-id:0;}
td
	{mso-style-parent:style0;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xl65
	{mso-style-parent:style0;
	font-size:9.0pt;
	text-align:center;
	border:.5pt solid #A5A5A5;
	background:#E5E0EC;
	mso-pattern:black none;}
.xl66
	{mso-style-parent:style0;
	font-size:9.0pt;
	text-align:center;
	border:.5pt solid #A5A5A5;}
-->
</style>
</head>

<body link=blue vlink=purple class=xl65>
<table border=0 cellpadding=0 cellspacing=0 width=1204 style='border-collapse:collapse;table-layout:fixed;width:905pt'>
 <col width=109 style='mso-width-source:userset;mso-width-alt:3488;width:82pt'>
 <col width=45 span=2 style='mso-width-source:userset;mso-width-alt:1440;width:34pt'>
 <col width=72 style='width:54pt'>
 <col width=45 span=2 style='mso-width-source:userset;mso-width-alt:1440;width:34pt'>
 <col width=104 span=2 style='mso-width-source:userset;mso-width-alt:3328;width:78pt'>
 <col width=72 style='width:54pt'>
 <col width=101 style='mso-width-source:userset;mso-width-alt:3232;width:76pt'>
 <col width=186 style='mso-width-source:userset;mso-width-alt:5952;width:140pt'>
 <col width=60 style='mso-width-source:userset;mso-width-alt:1920;width:45pt'>
 <col width=48 style='mso-width-source:userset;mso-width-alt:1536;width:36pt'>
 <col width=60 style='mso-width-source:userset;mso-width-alt:1920;width:45pt'>
 <col width=48 style='mso-width-source:userset;mso-width-alt:1536;width:36pt'>
 <col width=60 style='mso-width-source:userset;mso-width-alt:1920;width:45pt'>
 <tr height=22 style='height:16.5pt'>
  <td rowspan=2 height=44 class=xl65 width=109 style='height:33.0pt;width:82pt'>강사ID</td>
  <td rowspan=2 class=xl65 width=45 style='width:34pt'>강좌</td>
  <td rowspan=2 class=xl65 width=45 style='width:34pt'>마진</td>
  <td rowspan=2 class=xl65 width=72 style='width:54pt'>재료비마진</td>
  <td rowspan=2 class=xl65 width=45 style='width:34pt'>DIY</td>
  <td rowspan=2 class=xl65 width=45 style='width:34pt'>마진</td>
  <td rowspan=2 class=xl65 width=104 style='width:78pt'>스트리트명</td>
  <td rowspan=2 class=xl65 width=104 style='width:78pt'>회사명</td>
  <td rowspan=2 class=xl65 width=72 style='width:54pt'>담당자</td>
  <td rowspan=2 class=xl65 width=101 style='width:76pt'>전화번호</td>
  <td rowspan=2 class=xl65 width=186 style='width:140pt'>E-Mail / 등록일</td>
  <td colspan=2 class=xl65 width=108 style='border-left:none;width:81pt'>사용여부</td>
  <td colspan=3 class=xl65 width=168 style='border-left:none;width:126pt'>스트리트오픈여부</td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl65 style='height:16.5pt;border-top:none;border-left:
  none'>텐바이텐</td>
  <td class=xl65 style='border-top:none;border-left:none'>제휴몰</td>
  <td class=xl65 style='border-top:none;border-left:none'>텐바이텐</td>
  <td class=xl65 style='border-top:none;border-left:none'>제휴몰</td>
  <td class=xl65 style='border-top:none;border-left:none'>커뮤니티</td>
 </tr>
<% for i=0 to opartner.FresultCount-1 %>
 <tr height=44 style='height:24pt'>
  <td height=44 class=xl66 style='height:24pt;border-top:none'><%= opartner.FPartnerList(i).FID %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Flec_yn %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Flec_margin %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Fmat_margin %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Fdiy_yn %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Fdiy_margin %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).FSocName_Kor %><br style='mso-data-placement:same-cell;'><%= opartner.FPartnerList(i).FSocName %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Fcompany_name %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Fmanager_name %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Ftel %><br style='mso-data-placement:same-cell;'><%= opartner.FPartnerList(i).Fmanager_hp %></td>
  <td class=xl66 style='border-top:none;border-left:none'><%= opartner.FPartnerList(i).Femail %><br style='mso-data-placement:same-cell;'><%= opartner.FPartnerList(i).Fregdate %></td>
  <td class=xl66 style='border-top:none;border-left:none'><% if opartner.FPartnerList(i).Fisusing="Y" then %>O<% else %>X<% end if %></td>
  <td class=xl66 style='border-top:none;border-left:none'><% if opartner.FPartnerList(i).Fisextusing="Y" then %>O<% else %>X<% end if %></td>
  <td class=xl66 style='border-top:none;border-left:none'><% if opartner.FPartnerList(i).Fstreetusing="Y" then %>O<% else %>X<% end if %></td>
  <td class=xl66 style='border-top:none;border-left:none'><% if opartner.FPartnerList(i).Fextstreetusing="Y" then %>O<% else %>X<% end if %></td>
  <td class=xl66 style='border-top:none;border-left:none'><% if opartner.FPartnerList(i).Fspecialbrand="Y" then %>O<% else %>X<% end if %></td>
 </tr>
<% next %>
</table>
</body>
</html>
<%
set opartner = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->