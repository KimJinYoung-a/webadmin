<% 
'###########################################################
' Description :  월간 CS문의 및 클레임(~주) 그래프 용 xml 추출 (플래쉬 파일에 적용됨)
' History : 2007.08.24 한용민 생성
'###########################################################

class Cmonthcsclaimitem								'월간고객문의및 클래임용 
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
	
	public fyyyy				'년도
	public fmm					'달
	public fdd					'일
	
	public fitemd0				
	public fitemd1				
	public fitemd2				
	public fitemd3				
	public fitemd4				
	public fitemd5				
	public fitemd6						
end class

class Cchulgoitemlist
	public flist
	public frectyyyy
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
public sub fmonthcsclaim					'월간CS문의 및 클래임
		dim sql , i
	
		sql = sql & "select left(yyyymmdd,7) as yyyymm," 
		sql = sql & " count(case when divcd='a000' then left(yyyymmdd,7) end) as a000," 
		sql = sql & " count(case when divcd='a001' then left(yyyymmdd,7) end) as a001,"
		sql = sql & " count(case when divcd='a002' then left(yyyymmdd,7) end) as a002,"
		sql = sql & " count(case when divcd='a004' then left(yyyymmdd,7) end) as a004,"
		sql = sql & " count(case when divcd='a010' then left(yyyymmdd,7) end) as a010,"
		sql = sql & " count(case when divcd='a011' then left(yyyymmdd,7) end) as a011,"
		sql = sql & " count(case when divcd='a008' then left(yyyymmdd,7) end) as a008 "
		sql = sql & " from [db_datamart].[dbo].tbl_cs_daily_as_summary" 
		sql = sql & " where left(yyyymmdd,4) = '"& yyyy &"'" 
		sql = sql & " group by left(yyyymmdd,7)"
		
		rsget.open sql,dbget,1
		FTotalCount = rsget.recordcount
		redim flist(FTotalCount)
		i = 0
		
			if not rsget.eof then
				do until rsget.eof
					set flist(i) = new Cmonthcsclaimitem
						flist(i).fyyyy = rsget("yyyymm")		'날짜
						flist(i).fitemd0 = rsget("a000")		'맞교환출고
						flist(i).fitemd1 = rsget("a001")		'누락재발송
						flist(i).fitemd2 = rsget("a002")		'서비스발송
						flist(i).fitemd3 = rsget("a004")		'반품
						flist(i).fitemd4 = rsget("a010")		'회수
						flist(i).fitemd5 = rsget("a011")		'맞교환회수
						flist(i).fitemd6 = rsget("a008")		'주문취소
					rsget.movenext
					i = i + 1
				loop	
			end if
		rsget.close	
	end sub
end class
	

%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<% 
dim yyyy
yyyy = session("yyyy")

dim omonthcsclaim , i
	set omonthcsclaim = new Cchulgoitemlist
	omonthcsclaim.frectyyyy = yyyy
	omonthcsclaim.fmonthcsclaim()
%>
<% if omonthcsclaim.ftotalcount > 0 then %>	

	<?xml version='1.0' encoding='EUC-KR' ?>
	<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' lineThickness='3' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='2' plotBorderAlpha='20' >
	<categories>
		<% for i = 0 to omonthcsclaim.FTotalCount-1 %>
			<category name='<%= right(omonthcsclaim.flist(i).fyyyy,2) %>월' showName='1' showLine='1' />
		<% next %>	
	</categories>
	
	<dataset seriesName='맞교환출고' color='F60925' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd0 %>' />
		<% next %>
	
	</dataset>
	
	<dataset seriesName='누락재발송' color='F2F84A' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd1 %>' />
		<% next %>
	
	</dataset>
	<dataset seriesName='서비스발송' color='0611F9' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd2 %>' />
		<% next %>
	
	</dataset>
	<dataset seriesName='반품' color='94F84A' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd3 %>' />
		<% next %>
	
	</dataset>
	<dataset seriesName='회수' color='4E524B' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd4 %>' />
		<% next %>
	
	</dataset>
	<dataset seriesName='맞교환출고' color='865485' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd5 %>' />
		<% next %>
	
	</dataset>
	<dataset seriesName='주문취소' color='59F8BA' showValues='0' parentYAxis='P' >
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<set value='<%= omonthcsclaim.flist(i).fitemd6 %>' />
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
	
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->