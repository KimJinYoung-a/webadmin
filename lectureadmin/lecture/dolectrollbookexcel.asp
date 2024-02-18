<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
''통합 .2008-05-20
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, itemid, lecOption
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite

searchfield = RequestCheckvar(request("searchfield"),16)
userid = RequestCheckvar(request("userid"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
username = RequestCheckvar(request("username"),16)
userhp = RequestCheckvar(request("userhp"),16)
etcfield = RequestCheckvar(request("etcfield"),2)
etcstring = RequestCheckvar(request("etcstring"),32)
itemid = RequestCheckvar(request("itemid"),10)
lecOption   = RequestCheckvar(request("lecOption"),10)

checkYYYYMMDD = RequestCheckvar(request("checkYYYYMMDD"),1)
checkJumunDiv = RequestCheckvar(request("checkJumunDiv"),1)
checkJumunSite = RequestCheckvar(request("checkJumunSite"),1)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

jumundiv = RequestCheckvar(request("jumundiv"),16)
'==============================================================================
dim nowdate, searchnextdate

if (yyyy1="") then
        nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
'==============================================================================

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CLectureFingerOrder
ojumun.FPageSize = 200
ojumun.FCurrPage = page

if checkYYYYMMDD="Y" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if (checkJumunDiv = "Y") then
        if (jumundiv="flowers") then
        	ojumun.FRectIsFlower = "Y"
        elseif (jumundiv="lecture") then
                ojumun.FRectIsLecture = "Y"
        elseif (jumundiv="minus") then
                ojumun.FRectIsMinus = "Y"
        end if
end if

if (checkJumunSite = "Y") then
	ojumun.FRectExtSiteName = jumunsite
end if


if (searchfield = "orderserial") then
        '주문번호
        ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
        '고객아이디
        ojumun.FRectUserID = userid
elseif (searchfield = "username") then
        '구매자명
        ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
        '구매자핸드폰
        ojumun.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
        '기타조건
        if etcfield="01" then
        	ojumun.FRectBuyname = etcstring
        elseif etcfield="02" then
        	ojumun.FRectReqName = etcstring
        elseif etcfield="03" then
        	ojumun.FRectUserID = etcstring
        elseif etcfield="04" then
        	ojumun.FRectIpkumName = etcstring
        elseif etcfield="06" then
        	ojumun.FRectSubTotalPrice = etcstring
        elseif etcfield="07" then
        	ojumun.FRectBuyHp = etcstring
        elseif etcfield="08" then
        	ojumun.FRectReqHp = etcstring
        elseif etcfield="09" then
        	ojumun.FRectReqSongjangNo = etcstring
        end if
end if

if (searchfield = "itemid") then
	ojumun.FRectItemID = itemid
	ojumun.FREctItemOption=lecOption
	ojumun.FRectIsAvailJumun = "hidden"
	ojumun.GetFingerOrderListByItemID
else
	ojumun.GetFingerOrderList
end if

dim ix,i
dim totalavailcount


dim olecture
set olecture = new CLecture
olecture.FRectIdx = itemid

if (searchfield = "itemid") then
	olecture.GetOneLecture
end if

'// 옵션정보
dim oLectOption
Set oLectOption = New CLectOption
oLectOption.FRectidx = itemid
''oLectOption.FRectOptIsUsing = "Y"
if itemid<>"" then
	oLectOption.GetLectOptionInfo
end if

dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = itemid

if (searchfield = "itemid") then
	olecschedule.GetOneLecSchedule
end If

Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=LectureRollBook_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>출석부 파일</title>
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:.98in .39in .98in .39in;
	mso-header-margin:0in;
	mso-footer-margin:0in;
	mso-horizontal-page-align:center;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
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
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
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
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
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
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;}
.xl66
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;}
.xl67
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"Short Date";
	text-align:center;}
.xl68
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border:.5pt solid windowtext;
	background:#F2F2F2;
	mso-pattern:black none;
	white-space:normal;}
.xl69
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;}
.xl70
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border:.5pt solid windowtext;}
.xl71
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl72
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border:.5pt solid windowtext;}
.xl73
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:center;
	border:.5pt solid windowtext;
	white-space:normal;}
.xl74
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	text-align:center;
	border:.5pt solid windowtext;
	white-space:normal;}
-->
</style>
</head>

<body link=blue vlink=purple class=xl65>

<table border=0 cellpadding=0 cellspacing=0 width=1338 style='border-collapse: collapse;table-layout:fixed;width:1004pt'>
 <col class=xl65 width=105 style='mso-width-source:userset;mso-width-alt:2986; width:79pt'>
 <col class=xl65 width=97 span=3 style='mso-width-source:userset;mso-width-alt: 2759;width:73pt'>
 <col class=xl65 width=59 style='mso-width-source:userset;mso-width-alt:1678; width:44pt'>
 <col class=xl65 width=139 style='mso-width-source:userset;mso-width-alt:3953; width:104pt'>
 <col class=xl65 width=184 style='mso-width-source:userset;mso-width-alt:5233; width:138pt'>
 <col class=xl65 width=80 style='width:60pt'>
 <col class=xl65 width=80 span=6 style='width:60pt'>
 <tr height=16 style='height:12.0pt'>
  <td height=16 class=xl65 width=105 style='height:12.0pt;width:79pt'></td>
  <td class=xl65 width=97 style='width:73pt'></td>
  <td class=xl65 width=97 style='width:73pt'></td>
  <td class=xl65 width=97 style='width:73pt'></td>
  <td class=xl65 width=59 style='width:44pt'></td>
  <td class=xl65 width=139 style='width:104pt'></td>
  <td class=xl65 width=184 style='width:138pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
  <td class=xl65 width=80 style='width:60pt'></td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl68 width=105 style='height:19.5pt;width:79pt'>강좌명</td>
  <td colspan=6 class=xl72 style='border-left:none'><%= olecture.FOneItem.Flec_title %></td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl68 width=105 style='height:19.5pt;border-top:none;
  width:79pt'>강사명</td>
  <td colspan=6 class=xl72 style='border-left:none'><%= olecture.FOneItem.Flecturer_name %></td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl68 width=105 style='height:19.5pt;border-top:none;
  width:79pt'>수강료 / 재료비</td>
  <td colspan=6 class=xl73 width=673 style='border-left:none;width:505pt'><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %> / <% if olecture.FOneItem.Fmatinclude_yn="Y" then %>포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)<% else %>별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)<% end if %></td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl68 width=105 style='height:19.5pt;border-top:none;
  width:79pt'>강좌일시</td>
  <td colspan=6 class=xl72 style='border-left:none'><%= olecture.FOneItem.Flec_startday1 %> ~ <%= olecture.FOneItem.Flec_endday1 %> / <% if (olecture.FOneItem.Flec_startday1<>olecschedule.FItemList(0).Fstartdate) or (olecture.FOneItem.Flec_endday1<>olecschedule.FItemList(0).Fenddate) then %><%= olecschedule.FItemList(0).Fstartdate %> ~ <%= olecschedule.FItemList(0).Fenddate %><% end if %></td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=16 style='height:12.0pt'>
  <td height=16 class=xl66 style='height:12.0pt;border-top:none'>　</td>
  <td class=xl69>　</td>
  <td class=xl69>　</td>
  <td class=xl69>　</td>
  <td class=xl69>　</td>
  <td class=xl69>　</td>
  <td class=xl69>　</td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl68 width=105 style='height:19.5pt;border-top:none;
  width:79pt'>주문번호</td>
  <td class=xl68 width=97 style='border-top:none;border-left:none;width:73pt'>결제여부</td>
  <td class=xl68 width=97 style='border-top:none;border-left:none;width:73pt'>이름</td>
  <td class=xl68 width=97 style='border-top:none;border-left:none;width:73pt'>아이디</td>
  <td class=xl68 width=59 style='border-top:none;border-left:none;width:44pt'>인원</td>
  <td class=xl68 width=139 style='border-top:none;border-left:none;width:104pt'>연락처</td>
  <td class=xl68 width=184 style='border-top:none;border-left:none;width:138pt'>기타사항</td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
 <% if ojumun.FresultCount<1 then %>
 <% else %>
 <% for ix=0 to ojumun.FresultCount-1 %>
 <% totalavailcount = totalavailcount + ojumun.FItemList(ix).FItemNo %>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td height=26 class=xl70 style='height:19.5pt;border-top:none'><%= ojumun.FItemList(ix).FOrderSerial %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%= ojumun.FItemList(ix).IpkumDivName %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%= ojumun.FItemList(ix).Fentryname %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%= chrbyte(ojumun.FItemList(ix).FUserID,3,"Y") %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%= ojumun.FItemList(ix).FItemNo %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%= ojumun.FItemList(ix).Fentryhp %></td>
  <td class=xl70 style='border-top:none;border-left:none'><%'= ojumun.FItemList(ix).Fbeasongmemo %></td>
  <td colspan=5 class=xl65 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl67 style='mso-ignore:colspan'></td>
 </tr>
 <% next %>
 <% end if %>
 <tr height=26 style='mso-height-source:userset;height:19.5pt'>
  <td colspan=4 height=26 class=xl70 style='height:19.5pt'>　</td>
  <td class=xl70 style='border-top:none;border-left:none'><%= totalavailcount %></td>
  <td class=xl70 style='border-top:none;border-left:none'>　</td>
  <td class=xl70 style='border-top:none;border-left:none'>　</td>
  <td colspan=7 class=xl65 style='mso-ignore:colspan'></td>
 </tr>
</table>
</body>
</html>
<%
set olecture = Nothing
set olecschedule = Nothing
set oLectOption = Nothing
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->