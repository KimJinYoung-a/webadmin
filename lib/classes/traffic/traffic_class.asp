<%
'###########################################################
' Description :  �ٹ����� traffic analysis  Ŭ����
' History : 2007.09.04 �ѿ�� ����
'###########################################################

Class CtrafficOne
	
	public fyyyymmdd			'��¥
	public ftotalcount			'�湮�ڼ�
	public fpageview			'��������
	public fnewcount			'�űԹ湮�ڼ�
	public frecount				'��湮�ڼ�
	public frealcount			'�����湮�ڼ�
	
   Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class
'##################################################################
class Ctrafficlist							'�ٹ����� Ʈ���ȳ���  ��񿡼� ��������

	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectbuy_date
	public frectbuy_date1		
	
	public sub Ftrafficlist
	dim sql , i
	
	sql = "select yyyymmdd,totalcount,pageview,newcount,recount,realcount"
	sql = sql & " from db_datamart.dbo.tbl_traffic_analysis"
	sql = sql & " where 1=1 and yyyymmdd between '"&frectbuy_date&"' and '"&frectbuy_date1&"'"
	sql = sql & " order by yyyymmdd desc"
	
	'response.write sql			'������ �ѷ�����.
	db3_rsget.open sql,db3_dbget,1
	
	FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
	if not db3_rsget.eof then						'���ڵ��� ù��°�� �ƴ϶��
		do until db3_rsget.eof						'���ڵ��� ������ ���� ����
			set flist(i) = new CtrafficOne 			'Ŭ������ �ְ�
		
			flist(i).fyyyymmdd = db3_rsget("yyyymmdd")				'��¥
			flist(i).ftotalcount = db3_rsget("totalcount")			'�湮�ڼ�
			flist(i).fpageview = db3_rsget("pageview")				'��������
			flist(i).fnewcount = db3_rsget("newcount")				'�űԹ湮�ڼ�
			flist(i).frecount = db3_rsget("recount")				'��湮�ڼ�
			flist(i).frealcount = db3_rsget("realcount")			'�����湮�ڼ�
		
			db3_rsget.movenext
			i = i+1
			loop		
		end if
	db3_rsget.close
	end sub
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class
'##################################################################
class Ctrafficgraph							'�ٹ����� Ʈ���ȳ���  �׷�����

	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectyyyy					'���� �������� ���� ����
	public frectmm						'���� �������� ���� ����
	public function frecttot()			'����� ���� ������ �޾ƿͼ� ��ħ...
	frecttot = frectyyyy & frectmm
	end function
	public function frecttotnew()
	
		if frectyyyy <>"" and frectmm <> "" then												'��¥���� �ִٸ�
			frecttotnew = " and left(yyyymmdd,6) = "& frecttot &""		'���� ������ �˻� �ɼ��� ���δ�
		else 
			frecttotnew = " and left(yyyymmdd,6) = "& 0 &""				'�˻����� ���ٸ� �⺻��0�� �ְ�, �˻� �ɼ��� ���δ�.ùȭ�鿡�� ��絥���Ͱ� �ٻѷ����°��� ����...
		end if	
	end function
	
	public sub Ftrafficlist
	dim sql , i
	
	sql = "select yyyymmdd,totalcount,pageview,newcount,recount,realcount"
	sql = sql & " from db_datamart.dbo.tbl_traffic_analysis"
	sql = sql & " where 1=1 "& frecttotnew &""
	sql = sql & " order by yyyymmdd asc"
	
	'response.write sql			'������ �ѷ�����.
	db3_rsget.open sql,db3_dbget,1
	
	FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
	if not db3_rsget.eof then						'���ڵ��� ù��°�� �ƴ϶��
		do until db3_rsget.eof						'���ڵ��� ������ ���� ����
			set flist(i) = new CtrafficOne 			'Ŭ������ �ְ�
		
			flist(i).fyyyymmdd = db3_rsget("yyyymmdd")				'��¥
			flist(i).ftotalcount = db3_rsget("totalcount")			'�湮�ڼ�
			flist(i).fpageview = db3_rsget("pageview")				'��������
			flist(i).fnewcount = db3_rsget("newcount")				'�űԹ湮�ڼ�
			flist(i).frecount = db3_rsget("recount")				'��湮�ڼ�
			flist(i).frealcount = db3_rsget("realcount")			'�����湮�ڼ�
		
			db3_rsget.movenext
			i = i+1
			loop		
		end if
	db3_rsget.close
	end sub
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

%>
