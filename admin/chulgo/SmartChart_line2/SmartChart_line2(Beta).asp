<% 
'###########################################################
' Description :  월간 CS문의 및 클레임(~주) 그래프 용 xml 추출 (플래쉬 파일에 적용됨)
' History : 2007.08.23 한용민 생성
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
		sql = sql & " where left(yyyymmdd,4) = '"& frectyyyy &"'" 
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
	
	<?xml version="1.0" encoding="euc-kr"?>
	<setting>
	<!-- 제목 -->
			<text>월간CS</text>
	   		<textcolor>888888</textcolor>
	<!-- 타이틀 위치  -->
			<align>center</align>
	
	<!-- 카테고리칼라 (1~15까지 됨)-->
		<color>
			<groupcolor>1</groupcolor>
			<groupcolor>2</groupcolor>
			<groupcolor>3</groupcolor>
			<groupcolor>4</groupcolor>
			<groupcolor>5</groupcolor>
			<groupcolor>6</groupcolor>
			<groupcolor>7</groupcolor>
		</color>
	
	<!-- 카테고리 -->
		<category>
			<name>맞교환출고</name>
			<name>누락재발송</name>
			<name>서비스발송</name>
			<name>반품</name>
			<name>회수</name>
			<name>맞교환회수</name>
			<name>주문취소</name>
			
		</category>
	<!-- 카테고리 -->
			
			<Num>
				<% for i = 0 to omonthcsclaim.FTotalCount-1 %>
					<name><%= right(omonthcsclaim.flist(i).fyyyy,2) %>월</name>
				<% next %>
			</Num>
	  	
	  	<check_lengthwise>400</check_lengthwise>
	  	<!-- 그래프의 세로축 -->
	
	 	<check_crosswise>450</check_crosswise>
		<!-- 그래프의 가로축 -->
	
	  	<check_startx>50</check_startx>
	  	<!-- 그래프의 처음 시작 위치x값 -->
	 
	  	<check_starty>530</check_starty>
	  	<!-- 그래프의 처음 시작 위치y값 -->
	
	        <categoryX>50</categoryX>
	 	<!-- 카테고리 x값 -->
	
	        <categoryY>30</categoryY>
	  	<!-- 카테고리 y값 -->
	
	        <linecolor>3</linecolor>
	  	<!-- 스테이지 실선 색 설정 (1~15) -->
	
	        <linecolor>8</linecolor>
	  	<!-- 스테이지 점선 색 설정 (1~15) -->	
	</setting>
	<!-- 그룹0 -->
	
	<chart>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd0 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd1 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd2 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd3 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd4 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd5 %></value>
			<% next %>
	
		</group>
		<group>
			<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 	
			<value><%= omonthcsclaim.flist(i).fitemd6 %></value>
			<% next %>
	
		</group>
	</chart>
	
	</xml>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->