<%
'###########################################################
' Description :  �ٹ����� ȸ��Ż�� ��Ȳ
' History : 2008.02.15 �ѿ�� ����
'###########################################################

Class cuserwithdrawoneitem		'ȸ��������Ȳ
	public fwdrawDate		'��¥
	public fwdrawSex		'����
	public fwdrawAreaSido	'�ּ�(����)
	public fwdrawAreaGugun	'���ּ�
	public fwdrawAge		'����
	public fwdrawReason		'Ż�����
	public fwdrawCount		'Ż���
	public fwdrawReason_01	'��ǰǰ���Ҹ�
	public fwdrawReason_02	'�̿�󵵳���
	public fwdrawReason_03	'�������
	public fwdrawReason_04	'��������������
	public fwdrawReason_05	'��ȯ/ȯ��/ǰ���Ҹ�
	public fwdrawReason_06	'��Ÿ
	public fwdrawReason_07	'a/s�Ҹ�
	public fmancount		'���ڼ�
	public fgirlcount		'���ڼ�
	public fwithdrowtotalcount		
	
    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
end Class

class cuserwithdrawlist		
	public FItemList

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	
	public FRectStartdate	
	public FRectEndDate

	public sub fuserwithdrawlist()			'ȸ��Ż����Ȳ
		dim sqlstr, i
		
		sqlstr = "select convert(varchar(10),wdrawDate,121) as wdrawDate"
		sqlstr = sqlstr & " ,sum(wdrawCount) as withdrowtotalcount"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '' then wdrawCount end) as wdrawReason"		
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '01' then wdrawCount end) as wdrawReason_01"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '02' then wdrawCount end) as wdrawReason_02"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '03' then wdrawCount end) as wdrawReason_03"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '04' then wdrawCount end) as wdrawReason_04"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '05' then wdrawCount end) as wdrawReason_05"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '06' then wdrawCount end) as wdrawReason_06"
		sqlstr = sqlstr & " ,sum(case when wdrawReason = '07' then wdrawCount end) as wdrawReason_07"
		sqlstr = sqlstr & " ,sum(case when wdrawSex = '��' then wdrawCount end) as mancount"
		sqlstr = sqlstr & " ,sum(case when wdrawSex = '��' then wdrawCount end) as girlcount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by convert(varchar(10),wdrawDate,121)"
		sqlstr = sqlstr & " order by wdrawDate"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'���ڵ��� ù��°�� �ƴ϶��
			do until db3_rsget.eof						'���ڵ��� ������ ���� ����
				set FItemList(i) = new cuserwithdrawoneitem 			'Ŭ������ �ְ�
	
					FItemList(i).fwdrawDate = db3_rsget("wdrawDate")
					FItemList(i).fwithdrowtotalcount = db3_rsget("withdrowtotalcount")
					FItemList(i).fwdrawReason = db3_rsget("wdrawReason")
					FItemList(i).fwdrawReason_01 = db3_rsget("wdrawReason_01")
					FItemList(i).fwdrawReason_02 = db3_rsget("wdrawReason_02")				
					FItemList(i).fwdrawReason_03 = db3_rsget("wdrawReason_03")
					FItemList(i).fwdrawReason_04 = db3_rsget("wdrawReason_04")
					FItemList(i).fwdrawReason_05 = db3_rsget("wdrawReason_05")						
					FItemList(i).fwdrawReason_06 = db3_rsget("wdrawReason_06")
					FItemList(i).fwdrawReason_07 = db3_rsget("wdrawReason_07")
					FItemList(i).fmancount = db3_rsget("mancount")
					FItemList(i).fgirlcount = db3_rsget("girlcount")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub	

	public sub fuserwithdraw_sexgraph()			'ȸ��Ż����Ȳ(���� �׷�����)
		dim sqlstr, i
		
		sqlstr = "select"
		sqlstr = sqlstr & " wdrawSex,sum(wdrawCount) as wdrawCount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by wdrawSex"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'���ڵ��� ù��°�� �ƴ϶��
			do until db3_rsget.eof						'���ڵ��� ������ ���� ����
				set FItemList(i) = new cuserwithdrawoneitem 			'Ŭ������ �ְ�
	
					FItemList(i).fwdrawSex = db3_rsget("wdrawSex")
					FItemList(i).fwdrawCount = db3_rsget("wdrawCount")
					
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub		

	public sub fuserwithdraw_areagraph()			'ȸ��Ż����Ȳ(���� �׷�����)
		dim sqlstr, i
		
		sqlstr = "select"
		sqlstr = sqlstr & " wdrawReason,sum(wdrawCount) as wdrawCount"
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_withdraw_log"
		sqlstr = sqlstr & " where 1=1"
		
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),wdrawDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if			
		
		sqlstr = sqlstr & " group by wdrawReason"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then						'���ڵ��� ù��°�� �ƴ϶��
			do until db3_rsget.eof						'���ڵ��� ������ ���� ����
				set FItemList(i) = new cuserwithdrawoneitem 			'Ŭ������ �ְ�
	
					FItemList(i).fwdrawReason = db3_rsget("wdrawReason")
					FItemList(i).fwdrawCount = db3_rsget("wdrawCount")
					
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub	

    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub	
end class
%>