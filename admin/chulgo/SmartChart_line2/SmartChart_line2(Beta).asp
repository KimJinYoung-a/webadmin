<% 
'###########################################################
' Description :  ���� CS���� �� Ŭ����(~��) �׷��� �� xml ���� (�÷��� ���Ͽ� �����)
' History : 2007.08.23 �ѿ�� ����
'###########################################################

class Cmonthcsclaimitem								'���������ǹ� Ŭ���ӿ� 
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
	
	public fyyyy				'�⵵
	public fmm					'��
	public fdd					'��
	
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
	
public sub fmonthcsclaim					'����CS���� �� Ŭ����
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
						flist(i).fyyyy = rsget("yyyymm")		'��¥
						flist(i).fitemd0 = rsget("a000")		'�±�ȯ���
						flist(i).fitemd1 = rsget("a001")		'������߼�
						flist(i).fitemd2 = rsget("a002")		'���񽺹߼�
						flist(i).fitemd3 = rsget("a004")		'��ǰ
						flist(i).fitemd4 = rsget("a010")		'ȸ��
						flist(i).fitemd5 = rsget("a011")		'�±�ȯȸ��
						flist(i).fitemd6 = rsget("a008")		'�ֹ����
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
	<!-- ���� -->
			<text>����CS</text>
	   		<textcolor>888888</textcolor>
	<!-- Ÿ��Ʋ ��ġ  -->
			<align>center</align>
	
	<!-- ī�װ�Į�� (1~15���� ��)-->
		<color>
			<groupcolor>1</groupcolor>
			<groupcolor>2</groupcolor>
			<groupcolor>3</groupcolor>
			<groupcolor>4</groupcolor>
			<groupcolor>5</groupcolor>
			<groupcolor>6</groupcolor>
			<groupcolor>7</groupcolor>
		</color>
	
	<!-- ī�װ� -->
		<category>
			<name>�±�ȯ���</name>
			<name>������߼�</name>
			<name>���񽺹߼�</name>
			<name>��ǰ</name>
			<name>ȸ��</name>
			<name>�±�ȯȸ��</name>
			<name>�ֹ����</name>
			
		</category>
	<!-- ī�װ� -->
			
			<Num>
				<% for i = 0 to omonthcsclaim.FTotalCount-1 %>
					<name><%= right(omonthcsclaim.flist(i).fyyyy,2) %>��</name>
				<% next %>
			</Num>
	  	
	  	<check_lengthwise>400</check_lengthwise>
	  	<!-- �׷����� ������ -->
	
	 	<check_crosswise>450</check_crosswise>
		<!-- �׷����� ������ -->
	
	  	<check_startx>50</check_startx>
	  	<!-- �׷����� ó�� ���� ��ġx�� -->
	 
	  	<check_starty>530</check_starty>
	  	<!-- �׷����� ó�� ���� ��ġy�� -->
	
	        <categoryX>50</categoryX>
	 	<!-- ī�װ� x�� -->
	
	        <categoryY>30</categoryY>
	  	<!-- ī�װ� y�� -->
	
	        <linecolor>3</linecolor>
	  	<!-- �������� �Ǽ� �� ���� (1~15) -->
	
	        <linecolor>8</linecolor>
	  	<!-- �������� ���� �� ���� (1~15) -->	
	</setting>
	<!-- �׷�0 -->
	
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