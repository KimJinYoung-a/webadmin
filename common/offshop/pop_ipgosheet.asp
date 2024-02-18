<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ¿ÀÇÁ¶óÀÎ ÀÔÃâ°í ³»¿ª¼­
' History : 2009.04.07 ¼­µ¿¼® »ý¼º
'			2012.07.17 ÇÑ¿ë¹Î ¼öÁ¤
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idx ,oipchulmaster, oipchul ,obrand ,scheduleorexedate ,i
	idx = requestCheckVar(request("idx"),10)

if idx = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('idx °ªÀÌ ¾ø½À´Ï´Ù');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

set oipchulmaster = new CShopIpChul
	oipchulmaster.frect_IS_Maker_Upche = C_IS_Maker_Upche
	oipchulmaster.FRectIdx = idx
	oipchulmaster.GetIpChulMasterList

set oipchul = new CShopIpChul
	oipchul.FRectIdx = idx
	oipchul.GetIpChulDetail

if oipchulmaster.ftotalcount < 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('ÇØ´çµÇ´Â ÀÔÃâ°í³»¿ªÀÌ ¾ø½À´Ï´Ù');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

set obrand = new CBrandShopInfoItem
	obrand.FRectChargeId = oipchulmaster.FItemList(0).FChargeid
	obrand.GetBrandShopInFo

'if IsNULL(oipchulmaster.FItemList(0).FExecDt) then
	scheduleorexedate = replace(Left(CstR(oipchulmaster.FItemList(0).FScheduleDt),10),"-","/")
'else
'	scheduleorexedate = replace(Left(CstR(oipchulmaster.FItemList(0).FExecDt),10),"-","/")
'end if

if request("xl")<>"" then
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=" +oipchulmaster.FItemList(0).FChargeid + "_" + Left(CStr(now()),10) + "_" + CStr(oipchul.FRectIdx) + ".xls"
end if
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
<style id="SHOP¿ë ÀÔ°í³»¿ª¼­_29731_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font529731
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;}
.xl2429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl3029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl3329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\;\[Red\]\\-\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl3729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl3929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl4129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl4229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"Short Date";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\;\[Red\]\\-\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\;\[Red\]\\-\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl4929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl5029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl5529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:20.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:right;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\0022\\\0022\#\,\#\#0\;\[Red\]\\-\0022\\\0022\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl5929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:20.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"yyyy\\\/mm\\\/dd";
	text-align:left;
	vertical-align:bottom;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:bottom;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\0022\\\0022\#\,\#\#0\;\[Red\]\\-\0022\\\0022\#\,\#\#0";
	text-align:left;
	vertical-align:bottom;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:"\0022\\\0022\#\,\#\#0";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8229731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8329731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8429731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8529731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8629731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8729731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8829731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8929731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9029731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9129731
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:µ¸¿ò, monospace;
	mso-font-charset:129;
	mso-char-type:none;}
--></style>
</head>

<body>
<div id="SHOP¿ë ÀÔ°í³»¿ª¼­_29731" align=center x:publishsource="Excel">

<table x:str border=0 cellpadding=0 cellspacing=0 width=643 class=xl2629731
 style='border-collapse:collapse;table-layout:fixed;width:483pt'>
 <col class=xl2429731 width=97 style='mso-width-source:userset;mso-width-alt:
 2759;width:73pt'>
 <col class=xl2629731 width=213 style='mso-width-source:userset;mso-width-alt:
 6058;width:160pt'>
 <col class=xl2429731 width=18 style='mso-width-source:userset;mso-width-alt:
 512;width:14pt'>
 <col class=xl2629731 width=75 style='mso-width-source:userset;mso-width-alt:
 2133;width:56pt'>
 <col class=xl2629731 width=83 style='mso-width-source:userset;mso-width-alt:
 2360;width:62pt'>
 <col class=xl2629731 width=72 style='mso-width-source:userset;mso-width-alt:
 2048;width:54pt'>
 <col class=xl2629731 width=85 style='mso-width-source:userset;mso-width-alt:
 2417;width:64pt'>
 <tr height=35 style='height:26.25pt'>
  <td colspan=4 height=35 class=xl5929731 width=403 style='height:26.25pt;
  width:303pt'>ÀÔ°í³»¿ª¼­(<%= obrand.Fchargeid %>)</td>
  <td colspan=3 class=xl6429731 width=240 style='width:180pt' align='right'><%= oipchulmaster.FItemList(0).Fshopname %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl5529731 style='height:12.0pt'></td>
  <td class=xl5529731></td>
  <td class=xl5529731></td>
  <td class=xl5529731></td>
  <td class=xl5629731></td>
  <td class=xl2429731></td>
  <td class=xl5729731></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td rowspan=2 height=32 class=xl6029731 style='height:24.0pt'>³¯<span
  style="mso-spacerun: yes">&nbsp;&nbsp; </span>Â¥</td>
  <td rowspan=2 class=xl6529731 ><%= scheduleorexedate %></td>
  <td rowspan=6 class=xl6129731 width=18 style='border-bottom:1.0pt solid black;
  width:14pt'>°ø<br>
  <br>
  ±Þ<br>
  <br>
  ÀÚ</td>
  <td class=xl5029731>µî ·Ï<span style="mso-spacerun: yes">&nbsp; </span>¹ø È£</td>
  <td colspan=3 class=xl6629731 style='border-right:1.0pt solid black;
  border-left:none'><%= obrand.FSocNo %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl2829731 style='height:12.0pt;border-top:none'>»ó<span
  style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>È£</td>
  <td class=xl6929731 style='border-top:none;border-left:none'><%= obrand.FSocName %></td>
  <td class=xl2729731 style='border-top:none;border-left:none'>¼º<span
  style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>¸í</td>
  <td class=xl7029731 style='border-top:none;border-left:none'><%= obrand.FCeoName %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td rowspan=2 height=32 class=xl6029731 style='height:24.0pt'>»ó<span
  style="mso-spacerun: yes">&nbsp;&nbsp; </span>È£</td>
  <td rowspan=2 class=xl7129731><%= oipchulmaster.FItemList(0).Fshopname %></td>
  <td class=xl2829731 style='border-top:none'>»ç¾÷Àå¼ÒÀçÁö</td>
  <td colspan=3 class=xl7229731 style='border-right:1.0pt solid black;
  border-left:none'><%= obrand.FAddress %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl2829731 style='height:12.0pt;border-top:none'>¾÷<span
  style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>ÅÂ</td>
  <td colspan=3 class=xl7229731 style='border-right:1.0pt solid black;
  border-left:none'><%= obrand.FUptae %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td rowspan=2 height=32 class=xl6029731 style='height:24.0pt'>ÃÑ±Ý¾×</td>
  <td rowspan=2 class=xl7529731 >\<%= ForMatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
  <td class=xl2729731 style='border-top:none'>Á¾<span style="mso-spacerun:
  yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>¸ñ</td>
  <td colspan=3 class=xl7229731 style='border-right:1.0pt solid black;
  border-left:none'><%= obrand.FUpjong %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl5129731 style='height:12.0pt;border-top:none'>´ã<span
  style="mso-spacerun: yes">&nbsp;&nbsp; </span>´ç<span style="mso-spacerun:
  yes">&nbsp;&nbsp; </span>ÀÚ</td>
  <td class=xl7629731 style='border-top:none;border-left:none'>¡¡</td>
  <td class=xl5229731 style='border-top:none;border-left:none'>¿¬<span
  style="mso-spacerun: yes">&nbsp;&nbsp; </span>¶ô<span style="mso-spacerun:
  yes">&nbsp;&nbsp; </span>Ã³</td>
  <td class=xl7729731 style='border-top:none;border-left:none'>¡¡</td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl2529731 style='height:12.0pt'></td>
  <td class=xl7829731></td>
  <td class=xl2529731></td>
  <td class=xl4829731></td>
  <td class=xl4829731></td>
  <td class=xl4829731></td>
  <td class=xl7929731></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl2529731 style='height:12.0pt'></td>
  <td class=xl2529731></td>
  <td class=xl2429731></td>
  <td class=xl2629731></td>
  <td class=xl2629731></td>
  <td class=xl2629731></td>
  <td class=xl4529731></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl3729731 style='height:12.75pt'>»óÇ°ÄÚµå</td>
  <td class=xl3829731>»óÇ°¸í</td>
  <td colspan=2 class=xl8029731 style='border-right:.5pt solid black;
  border-left:none'>¿É¼Ç</td>
  <td class=xl3929731>¼ÒºñÀÚ°¡</td>
  <td class=xl3829731>¼ö·®</td>
  <td class=xl3429731>¼ÒºñÀÚ°¡°è</td>
 </tr>
 <% for i=0 to oipchul.FResultCount -1 %>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl4929731 width=97 style='height:12.75pt;border-top:none;
  width:73pt'><%= oipchul.FItemList(i).FItemGubun %>-<%= CHKIIF(oipchul.FItemList(i).FShopItemId>=1000000,Format00(8,oipchul.FItemList(i).FShopItemId),Format00(6,oipchul.FItemList(i).FShopItemId)) %>-<%= oipchul.FItemList(i).FItemOption %></td>
  <td class=xl4029731 width=213 style='border-top:none;width:160pt'>
  	<%= oipchul.FItemList(i).FItemName %>
  </td>
  <td colspan=2 class=xl8229731 width=93 style='border-right:.5pt solid black;
  border-left:none;width:70pt'><%= oipchul.FItemList(i).FItemOptionName %></td>
  <td class=xl4129731 width=83 style='border-top:none;width:62pt'><%= FormatNumber(oipchul.FItemList(i).FSellcash,0) %></td>
  <td class=xl3529731 style='border-top:none'><%= oipchul.FItemList(i).FItemNo %></td>
  <td class=xl4629731 style='border-top:none'><%= FormatNumber(oipchul.FItemList(i).FItemNo*oipchul.FItemList(i).FSellcash,0) %></td>
 </tr>
 <% next %>
 <tr height=22 style='mso-height-source:userset;height:16.5pt'>
  <td height=22 class=xl4229731 style='height:16.5pt'>ºñ°í</td>
  <td class=xl5329731>¡¡</td>
  <td colspan=2 class=xl5329731 style='border-right:.5pt solid black'>¡¡</td>
  <td class=xl5429731 width=83 style='border-left:none;width:62pt'>¼ÒºñÀÚ°¡°è</td>
  <td class=xl5429731 width=72 style='border-left:none;width:54pt'>¡¡</td>
  <td class=xl5829731 style='border-left:none' >\<%= ForMatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl2429731 style='height:12.0pt'></td>
  <td class=xl2629731></td>
  <td colspan=2 class=xl4429731>¡¡</td>
  <td class=xl3129731></td>
  <td class=xl2429731></td>
  <td class=xl3329731></td>
 </tr>
 <tr height=30 style='mso-height-source:userset;height:22.5pt'>
  <td height=30 class=xl4229731 style='height:22.5pt'>ÀÎ°èÀÚ</td>
  <td class=xl8829731 style='border-left:none'>(ÀÎ)</td>
  <td colspan=2 class=xl9029731 style='border-right:1.0pt solid black'>ÀÎ¼öÀÚ</td>
  <td class=xl4329731 style='border-left:none'>¡¡</td>
  <td class=xl4429731>¡¡</td>
  <td class=xl9129731>(ÀÎ)</td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=97 style='width:73pt'></td>
  <td width=213 style='width:160pt'></td>
  <td width=18 style='width:14pt'></td>
  <td width=75 style='width:56pt'></td>
  <td width=83 style='width:62pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=85 style='width:64pt'></td>
 </tr>
 <![endif]>
</table>

</div>
</body>
</html>

<%
set obrand = Nothing
set oipchulmaster = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->