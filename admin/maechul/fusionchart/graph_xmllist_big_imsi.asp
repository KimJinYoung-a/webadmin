<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
response.charset = "utf-8"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<% 
Dim vGubun
dim Omaechul_list

dim yyyy1, yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i, mm1, mm2, defaultdate1, monthday
dim ipkumdatesucc, vParam
	yyyy1 			= request("yyyy1")
	yyyy2 			= request("yyyy2")
	dateview1 		= request("dateview1")
	datecancle 		= request("datecancle")
	bancancle 		= request("bancancle") 
	accountdiv 		= request("accountdiv")			
	sitename 		= request("sitename") 
	ipkumdatesucc 	= request("ipkumdatesucc")
	mm1 			= request("mm1")
	mm2 			= request("mm2")
	monthday		= request("monthday")
	
	vParam = request("param")

	yyyy1 = split(vParam,"^^")(1)
	yyyy1 = split(yyyy1,"=")(1)
	
	yyyy2 = split(vParam,"^^")(2)
	yyyy2 = split(yyyy2,"=")(1)

	datecancle = split(vParam,"^^")(3)
	datecancle = split(datecancle,"=")(1)
	
	bancancle = split(vParam,"^^")(4)
	bancancle = split(bancancle,"=")(1)
	
	accountdiv = split(vParam,"^^")(5)
	accountdiv = split(accountdiv,"=")(1)
	
	sitename = split(vParam,"^^")(6)
	sitename = split(sitename,"=")(1)
	
	dateview1 = split(vParam,"^^")(7)
	dateview1 = split(dateview1,"=")(1)
	
	ipkumdatesucc = split(vParam,"^^")(8)
	ipkumdatesucc = split(ipkumdatesucc,"=")(1)
	
	mm1 = split(vParam,"^^")(9)
	mm1 = split(mm1,"=")(1)
	
	mm2 = split(vParam,"^^")(10)
	mm2 = split(mm2,"=")(1)
	
	monthday = split(vParam,"^^")(10)
	monthday = split(monthday,"=")(1)
	
	
	defaultdate1 = dateadd("d",-60,year(now) & "-" &TwoNumber(month(now)) & "-" & day(now))		'날짜값이 없을때 기본값으로 60이전까지 검색
	if yyyy2 = "" then yyyy2 = year(now)
	if yyyy1 = "" then yyyy1 = CInt(yyyy2)-2
	if mm1 = "" then mm1 = "01"
	if mm2 = "" then mm2 = month(now)
	mm2 = TwoNumber(mm2)
	if bancancle = "" then bancancle = "1"
	if dateview1 = "" then dateview1 = "yes"

	'<!-- //-->


	set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy2
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc		
	Omaechul_list.fmaechul_graph()
	
%>

<?xml version='1.0' encoding='UTF-8' ?>
<chart chartBottomMargin='2' formatNumberScale='0' drawAnchors='1' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='5' plotBorderAlpha='20' >
<categories>
<%	for i = 0 to Omaechul_list.ftotalcount - 1	%>
	<category name='<%=mid(Omaechul_list.flist(i).forderdate,6,2)%>월' showName='1' showLine='1' />
<%	next	%>
</categories>	
<dataset seriesName='실금액' showValues='0' parentYAxis='P'>
<%	for i=0 to Omaechul_list.FTotalCount - 1	%>
	<set value='<%=Omaechul_list.flist(i).fsubtotalprice%>' />
<%	next	%>
</dataset>
	
<dataset seriesName='순수익' showValues='0' parentYAxis='P'>
<%	for i=0 to Omaechul_list.FTotalCount - 1	%>
	<set value='<%=Omaechul_list.flist(i).fsunsuik%>' />
<%	next	%>
</dataset>
	
<dataset seriesName='총건수' renderas="line" parentyaxis="S" showvalues="0">
<%	for i=0 to Omaechul_list.FTotalCount - 1	%>
	<set value='<%=Omaechul_list.flist(i).ftotalcount%>' />
<%	next	%>
</dataset>
	
</chart>

<%
	set Omaechul_list = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
