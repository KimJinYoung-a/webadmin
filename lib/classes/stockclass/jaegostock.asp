<%

'###########################################################
' Description :  ����ľ�
' History : 2007.07.13 �ѿ�� ����
' History : 2007.11.28 �ѿ�� ����
'###########################################################

class Cfitem					'Ŭ���� ����
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx				'�ε�����ȣ
	public fitemgubun		'��ǰ����
	public fitemid			'��ǰ��ȣ
	public fitemoption		'�ɼ��ڵ�	
	public fitemname		'��ǰ��
	public fitemoptionname	'�ɼǸ�
	public fmakerid			'�귣��id
	public fregdate			'�����
	public freguserid		'������id	
	public factionusername	'����ľ��ѻ��
	public factionstartdate	'����ľ��Ͻ�
	public fbasicstock		'����ľ����
	public frealstock		'����ľ� �ǻ簹��
	public ferrstock		'����
	public ffinishuserid	'�Ϸ���id
	public fstatecd			'�����ڵ�
	public deleteyn			'��������
	public makerid			'�˻����ʿ��Ѻ귣��id
	public fstats			'����
	public fsmallimage		'��ǰ�̹���
	public fbigo			'���
	public foptioncnt		'�ɼǺ񱳽� �ʿ��� ����
	public frealstocks		'realstock + offconfirmno + ipkumdiv5
	
	public function getbigoName()
		if fbigo = 1 then
		 	getbigoName = "�۾�����"
		elseif fbigo = 5 then
			getbigoName = "����ľǿϷ�"
		elseif fbigo = 7 then
			getbigoName = "�Ϸ�(�ݿ���)"
		elseif fbigo = 8 then
			getbigoName = "�Ϸ�(�̹ݿ�)"
		else
			getbigoName =""
		end if
	end function
end class

class Cfitemlist					'Ŭ���� ����
	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public Frectidx					'�ε��� ���� �ޱ� ���� ����
	public Frectitemid				'��ǰid ���� �ޱ� ���� ����
	public frectmakerid				'�귣�� ���� �ޱ� ���� ����
	public frectstats				'���� ���� �ޱ� ���� ���� 
	public frectorderingdate		'�۾��������� �ޱ� ���� ����
	public frectguestlist			'�Ϲݻ���ڿ� ����Ʈ�� �Ѹ��� ���� ����
	public frectitemoption
	
	public Sub fjaegoinsert()			'����Է¿�
		dim sqlStr ,i 
		
		sqlStr = "select" 
		sqlstr = sqlstr & " isnull(b.realstock,'0') as realstock,"
		sqlstr = sqlstr & " a.itemid, isnull(c.itemoption,'0000') as itemoption," 
		sqlstr = sqlstr & " a.itemgubun,isnull(a.itemgubun,'10') as itemgubun,"
		sqlstr = sqlstr & " a.smallimage,a.makerid ,a.itemname"
		sqlstr = sqlstr & " from db_item.[dbo].tbl_item a"
		sqlstr = sqlstr & " left join [db_summary].dbo.tbl_current_logisstock_summary b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on b.itemid = c.itemid and b.itemoption = c.itemoption" 
		sqlstr = sqlstr & " where 1=1 and a.itemid = '" & frectitemid &"'"
			if frectitemoption <> "0000" then 		
				sqlstr = sqlstr & " and c.itemoption = '" & frectitemoption &"'"
			end if
		rsget.Open sqlStr,dbget,1					'����ľ��������̺�(f)�� ������̺�(r)�� �����ؼ� ��ǰ ����Ʈ�� �����´�.
		'response.write sqlstr&"<br>"
				   	
	   	FTotalCount = rsget.recordcount
	   	redim flist(FTotalCount)
		i=0
			
		if  not rsget.EOF  then
			do until rsget.eof
				set flist(i) = new Cfitem
				
						
				flist(i).fitemgubun = rsget("itemgubun")				'��ǰ����
				flist(i).fitemid = rsget("itemid")						'��ǰ��ȣ
				flist(i).fitemoption = rsget("itemoption")				'�ɼ��ڵ�	
				flist(i).fitemname = rsget("itemname")					'��ǰ��
				flist(i).fmakerid = rsget("makerid")					'�귣��id
				flist(i).frealstock = rsget("realstock")				'����ľ� �ǻ簹��
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'��ǰ�̹���
				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub	
	
	public Sub fprintlist()			' ����Ʈ ��� �κ�
		dim sqlStr ,i

		sqlStr = "select" 
		sqlstr = sqlstr & " a.idx, b.smallimage, a.itemid, b.makerid ,a.basicstock," 
		sqlstr = sqlstr & " b.itemname, a.itemoption, c.optionname, a.statecd, a.actiondate," 
		sqlstr = sqlstr & " isnull(d.realstock,'0') as realstock" 
		sqlstr = sqlstr & " from [db_summary].[dbo].tbl_req_realstock a" 
		sqlstr = sqlstr & " join db_item.[dbo].tbl_item b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on a.itemid = c.itemid and a.itemoption = c.itemoption" 
		sqlstr = sqlstr & " left join db_summary.dbo.tbl_current_logisstock_summary d" 
		sqlstr = sqlstr & " on a.itemid = d.itemid and a.itemoption = d.itemoption" 			 		
		sqlstr = sqlstr & " Where a.idx in (" + Frectidx + ")"
		sqlstr = sqlstr & " order by idx desc"					
		rsget.Open sqlStr,dbget,1				'����ľ��������̺�(f)�� ������̺�(r)���� ��ǰid�� ��ǰ�ɼ��� ���� ���� �����´�.
		'response.write sqlstr&"<br>"
					   	
	   	FTotalCount = rsget.recordcount
	   	redim flist(FTotalCount)
		i=0
			
		if  not rsget.EOF  then
			do until rsget.eof
				set flist(i) = new Cfitem								'Ŭ������ �ְ�
				
				flist(i).fidx = rsget("idx")		
				flist(i).fitemid = rsget("itemid")						'��ǰ��ȣ
				flist(i).fitemoption = rsget("itemoption")				'�ɼ��ڵ�	
				flist(i).fitemname = rsget("itemname")					'��ǰ��
				flist(i).fitemoptionname = rsget("optionname")		'�ɼǸ�
				flist(i).fmakerid = rsget("makerid")					'�귣��id
				flist(i).factionstartdate = rsget("actiondate")			'����ľ��Ͻ�
				flist(i).fbasicstock = rsget("basicstock")				'����ľ����
				flist(i).frealstock = rsget("realstock")				'����ľ� �ǻ簹��
				flist(i).fstatecd = rsget("statecd")					'�����ڵ�
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'��ǰ�̹���
				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub	
	
	public Sub fwritelist()				'��ǰ �ɼǺ��� �˻�
		dim sql554 ,i 
			
			sql554 = "select o.optionname,isnull(o.itemoption,'0000') as itemoption,r.optioncnt from [db_item].[dbo].tbl_item r" 
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_option o on r.itemid = o.itemid" 
			sql554 = sql554 & " where r.itemid = '"& Frectitemid &"'"
			rsget.open sql554,dbget,1			'��ǰ���̺�(r)�� ��ǰ�ɼ����̺�(o)�� ���ؼ� ��ǰ�� ���������� �����´�.
			'response.write sql554&"<br>"
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem							'Ŭ�����ְ�
					flist(i).fitemoptionname = rsget("optionname")		'��ǰ�ɼ��̸��ְ�	
					flist(i).fitemoption = rsget("itemoption")			'��ǰ�ɼ��ڵ�ְ�
					flist(i).foptioncnt = rsget("optioncnt")			'�� ��ǰ�� �� �ɼ��� ����� �� �ְ�
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub
	
	public Sub fjonglist()						'��ǰ ����Ʈ �ѷ��ִ� �κ�
		dim sql , i , sqlcount , cnt
		
		sqlcount = "select count(idx) as cnt from [db_summary].[dbo].tbl_req_realstock"		'�˻� ���ð��� �ش��ϴ� �ε������� �����´�
		sqlcount = sqlcount & " where 1=1"
		'response.write sqlcount
		
		if Frectguestlist <> "" then
			sqlcount = sqlcount & " and statecd in (" & Frectguestlist & ")"
		end if
		
		if frectstats <> "" then												'������ request ���� ���� ���� �ִٸ�
			sqlcount = sqlcount & " and statecd = " & frectstats & ""			'���� ������ �˻� �ɼ��� ���δ�
		end if
		
		rsget.open sqlcount,dbget,1
		FTotalCount = rsget("cnt")				'�ѷ��ڵ� ���� �ε���ī��Ʈ�� �ְ�
		rsget.close
		
		sql = "select top "& FPageSize*FCurrpage &""
		sql = sql & " (isnull(d.realstock,'0')+isnull(d.offconfirmno,'0')+isnull(d.ipkumdiv5,'0')) as realstocks,"		
		sql = sql & " b.smallimage,b.itemname,b.makerid,b.smallimage,"
		sql = sql & " c.optionname , a.*"
		sql = sql & " from [db_summary].[dbo].tbl_req_realstock a"
		sql = sql & " join db_item.[dbo].tbl_item b"
		sql = sql & " on a.itemid = b.itemid"
		sql = sql & " left join [db_item].[dbo].tbl_item_option c" 
		sql = sql & " on a.itemid = c.itemid and a.itemoption = c.itemoption"
		sql = sql & " left join db_summary.dbo.tbl_current_logisstock_summary d"
		sql = sql & " on a.itemid = d.itemid and a.itemoption = d.itemoption"				
		sql = sql & " where 1=1"
	
		if Frectguestlist <> "" then
			sql = sql & " and statecd in (" & Frectguestlist & ")"
		end if	
		
		if frectmakerid <> "" then									'������ request ���� ����Ŀ ���� �ִٸ�
			sql = sql & " and makerid = '" & frectmakerid & "'"		'���� ������ �˻� �ɼ��� ���δ�
		end if 
		
		if frectstats <> "" then								'������ request ���� ���� ���� �ִٸ�
			sql = sql & " and statecd = " & frectstats & ""		'���� ������ �˻� �ɼ��� ���δ�
		end if

		sql = sql & " order by idx desc"
		'response.write sql			'���� �ѷ�����
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		
		FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1	
		
		redim flist(FResultCount)
		i=0			'������ i ���� o�ְ�
			
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set flist(i) = new Cfitem		'Ŭ�����ְ�
				
				flist(i).fidx = rsget("idx")		
				flist(i).fitemgubun = rsget("itemgubun")				'��ǰ����
				flist(i).fitemid = rsget("itemid")						'��ǰ��ȣ
				flist(i).fitemoption = rsget("itemoption")				'�ɼ��ڵ�	
				flist(i).fitemname = rsget("itemname")					'��ǰ��
				flist(i).fitemoptionname = rsget("optionname")		'�ɼǸ�
				flist(i).fmakerid = rsget("makerid")					'�귣��id
				flist(i).fregdate = rsget("regdate")					'�����
				flist(i).freguserid = rsget("reguserid")				'������id	
				flist(i).factionstartdate = rsget("actiondate")			'����ľ��Ͻ�
				flist(i).fbasicstock = rsget("basicstock")				'����ľ����
				flist(i).frealstock = rsget("realstock")				'����ľ� �ǻ簹��
				flist(i).ferrstock = rsget("errstock")					'����
				flist(i).ffinishuserid = rsget("finishuserid")			'�Ϸ���id
				flist(i).fstatecd = rsget("statecd")					'�����ڵ�
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				'��ǰ�̹���
				flist(i).fbigo = rsget("statecd")						'������ �Էµ���� �� ���°�
				flist(i).frealstocks = rsget("realstocks")						'������ �Էµ���� �� ���°�				
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub
	
	
	public Sub fbrandinsert()			'�귣�� ��� �˻� ������ ���� Ŭ����
		dim sqlstr ,i 

		sqlstr = "select" 
		sqlstr = sqlstr & " isnull(b.realstock,'0') as realstock,"
		sqlstr = sqlstr & " a.itemid , a.makerid , a.smallimage,a.itemname"
		sqlstr = sqlstr & " ,isnull(c.itemoption,'0000') as itemoption,"
		sqlstr = sqlstr & " c.optionname"
		sqlstr = sqlstr & " from db_item.[dbo].tbl_item a"
		sqlstr = sqlstr & " join [db_summary].dbo.tbl_current_logisstock_summary b" 
		sqlstr = sqlstr & " on a.itemid = b.itemid" 
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option c" 
		sqlstr = sqlstr & " on b.itemid = c.itemid and b.itemoption = c.itemoption" 
		sqlstr = sqlstr & " where 1=1"
		
		if frectmakerid <> "" then
			sqlstr = sqlstr & " and a.makerid= '" & frectmakerid &"'"
		end if 
		if Frectitemid <> "" then
			sqlstr = sqlstr & " and a.itemid= '" & Frectitemid &"'"
		end if 
	
		sqlstr = sqlstr & " order by a.itemid desc"
		
		rsget.Open sqlstr,dbget,1					'
		'response.write sqlstr&"<br>"	   	
	   	FTotalCount = rsget.recordcount
		redim flist(FTotalCount)
		i=0			'������ i ���� o�ְ�
			
		if not rsget.EOF then
			do until rsget.eof
				set flist(i) = new Cfitem		'Ŭ�����ְ�
				
				flist(i).fitemid = rsget("itemid")		
				flist(i).fmakerid = rsget("makerid")				
				flist(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")			
				flist(i).fitemname = rsget("itemname")					
				flist(i).fitemoption = rsget("itemoption")		
				flist(i).fitemoptionname = rsget("optionname")					
				flist(i).fbasicstock = rsget("realstock")					
						
				rsget.moveNext		   
				i=i+1
				
			loop
		end if
		rsget.close
	end Sub
			
	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 11
		FResultCount = 0
		FScrollCount = 11
		FTotalCount =0
	end sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1								'//���� �������� 1���� ũ�� ����
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1	'//��ü �������� ����������+��ü��������ũ��-1�� ������ ũ�� ����
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1	'//���� �������� �������������� 1�� ���� ��ü��������ũ���� ������ ��ü��������ũ���� ������ +1�� �ϸ� ����. 
	end Function
end class

'// ��ǰ �̹��� ��θ� ����Ͽ� ��ȯ //
function GetImageSubFolderByItemid(byval iitemid)
    if (iitemid <> "") then
	    GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	else
	    GetImageSubFolderByItemid = ""
	end if
end function
%>