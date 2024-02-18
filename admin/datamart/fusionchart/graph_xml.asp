<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/baesongtermCls.asp" -->

<% 
dim Obaesong_list

Dim i, vTop, vSDate, vEDate, vGubun, vParam, vItemID, vMakerID

	vParam = request("param")

	vGubun = split(vParam,"^^")(1)
	vGubun = split(vGubun,"=")(1)

	vSDate = split(vParam,"^^")(2)
	vSDate = split(vSDate,"=")(1)
	
	vEDate = split(vParam,"^^")(3)
	vEDate = split(vEDate,"=")(1)
	
	vItemID = split(vParam,"^^")(4)
	vItemID = split(vItemID,"=")(1)
	
	vMakerID = split(vParam,"^^")(5)
	vMakerID = split(vMakerID,"=")(1)


	'<!-- //-->


	set Obaesong_list = new Cbaesong_list
	Obaesong_list.FSDate = vSDate
	Obaesong_list.FEDate = vEDate
	Obaesong_list.FGubun = vGubun
	Obaesong_list.FIsNotZero = "Y"
	Obaesong_list.FItemID = vItemID
	'Obaesong_list.FMakerID = vMakerID
	Obaesong_list.fbaesong_graph()
	

%>
	<?xml version='1.0' encoding='EUC-KR' ?>
	<chart bgColor='f1f1f1' canvasbgcolor='BDBD00,FFFFFF' canvasbgalpha='60' canvasbgangle='270' outcnvBaseFontColor='000000' caption="" subCaption="" yaxisname="硅价家夸老" xaxisname="" alternateHGridAlpha="30" alternateHGridColor="FFFFFF" canvasBorderThickness="1" canvasBorderColor="114B78" hoverCapBorderColor="114B78" hoverCapBgColor="E7EFF6" plotGradientColor="DCE6F9" plotFillAngle='90' plotFillColor='000000' plotfillalpha='80' showAnchors='0' canvaspadding='20' plotFillRatio='10,90' showPlotBorder='1' plotBorderColor='FFFFFF' plotBorderAlpha='20' divlinecolor='FFFFFF' divlinealpha='60' numberSuffix='老'>
		<%	for i = 0 to Obaesong_list.ftotalcount - 1	%>
		<set label="<%=Obaesong_list.flist(i).fyyyy%>-<%=TwoNumber(Obaesong_list.flist(i).fmm)%>" value="<%=Round(Obaesong_list.flist(i).fdelayday,2)%>" />
		<%	next	%>
		<styles>
			<definition>
				<style name='DataValuesFont' type='font' color='000000' />
				<style name='myAnim' type='animation' param='_alpha' start='0' duration='1' />
				<style name='dummyShadow' type='Shadow' alpha='1' />
			</definition>
			<application>
				<apply toObject='DATAVALUES' styles='DataValuesFont,dummyShadow,myAnim' />
			</application>	
		</styles>
	</chart>
<%
	set Obaesong_list = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

