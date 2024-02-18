<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일 오픈율 관리
' History : 2007.08.27 한용민 생성
' History : 2012.12.04 전면 수정
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mailopenresultclass/mailopenclass2.asp"-->
<%
Dim yyyy , mm, gubun
yyyy = session("yyyy")
mm = session("mm")
gubun = request("gubun")

Dim oMailzine , i
Set oMailzine = new CMailzinelistgraph
	oMailzine.FRectyyyy = yyyy
	oMailzine.FRectmm = mm
	oMailzine.fGubun = gubun
	oMailzine.FMailzinelist()
%>
<?xml version='1.0' encoding='EUC-KR' ?>
<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' lineThickness='3' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='2' plotBorderAlpha='20' >
	<categories>
	<% For i = 0 to oMailzine.FTotalCount-1 %>
		<category name='<%= right(oMailzine.flist(i).Freenddate,2) %>일' showName='1' showLine='1' />
	<% Next %>	
	</categories>
	
	<dataset seriesName='오픈률(%)' color='F60925' showValues='0' parentYAxis='S' >
	<% For i=0 to oMailzine.FTotalCount - 1 %> 	
		<set value='<%= oMailzine.flist(i).Fopenpct %>' />
	<% Next %>
	</dataset>
	
	<dataset seriesName='성공발송률(%)' color='F2F84A' showValues='0' parentYAxis='P' >
	<% For i=0 to oMailzine.FTotalCount - 1 %> 	
		<set value='<%= oMailzine.flist(i).Fsuccesspct %>' />
	<% Next %>
	</dataset>
	<dataset seriesName='실제발송률(%)' color='0611F9' showValues='0' parentYAxis='P' >
	<% For i=0 to oMailzine.FTotalCount - 1 %> 	
		<set value='<%= oMailzine.flist(i).Frealpct %>' />
	<% Next %>
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
<% Set oMailzine = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->