<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ÅÙ¹ÙÀÌÅÙ traffic analysis  
' History : 2007.09.04 ÇÑ¿ë¹Î »ý¼º
'###########################################################
%>

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<% 
dim yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i
dim ipkumdatesucc					
	yyyy2 = session("yyyy2")
	dateview1 = session("dateview1")
	datecancle = session("datecancle")
	bancancle = session("bancancle") 
	accountdiv = session("accountdiv")			
	sitename = session("sitename") 
	ipkumdatesucc = session("ipkumdatesucc")

dim Omaechul_list
set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy2	
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc		
	Omaechul_list.fmaechul_graph()
%>



	<?xml version='1.0' encoding='EUC-KR' ?>
	<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' lineThickness='3' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='2' plotBorderAlpha='20' >
	<categories>
		<% for i = 0 to Omaechul_list.ftotalcount - 1 %>
			<category name='<%= mid(Omaechul_list.flist(i).forderdate,6,2) %>¿ù' showName='1' showLine='1' />
		<% next %>	
	</categories>
	
	<dataset seriesName='½Ç±Ý¾×' color='F60925' showValues='0' parentYAxis='P' >
		<% for i=0 to Omaechul_list.FTotalCount - 1 %> 	
			<set value='<%= Omaechul_list.flist(i).fsubtotalprice %>' />
		<% next %>	
	</dataset>
	
	<dataset seriesName='¼ø¼öÀÍ' color='F2F84A' showValues='0' parentYAxis='P' >
		<% for i=0 to Omaechul_list.FTotalCount - 1 %> 	
			<set value='<%= Omaechul_list.flist(i).fsunsuik %>' />
		<% next %>	
	</dataset>
	<dataset seriesName='ÃÑ°Ç¼ö' color='0611F9' showValues='0' parentYAxis='S' >
		<% for i=0 to Omaechul_list.FTotalCount - 1 %> 	
			<set value='<%= Omaechul_list.flist(i).ftotalcount %>' />
		<% next %>	
	</dataset>

	<trendLines></trendLines>
	<styles>
		<definition>
			<style name='shadow215' type='shadow' angle='215' distance='3'/>
			<style name='shadow45' type='shadow' angle='45' distance='3'/>
		</definition>
		<application>
			<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />
			<apply toObject='DATAPLOTLINE' styles='shadow215' />
			<apply toObject='DATAPLOT' styles='shadow215' />
		</application>
	</styles>
	</chart>
	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
