<%
'###########################################################
' Description :  ���ϸ��� ���� 
' History : 2007.10.23 �ѿ�� ����
'###########################################################

Class Cmileageoneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fjukyocd
	public fjukyoname
	public fisusing
	
end class

class Cmileagelist
	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 0
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0	
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public frectjukyocd
	public frectisusing
	public frectseachjukyocd
	
'##############################################################################################		
	public sub fmileagelist									
	dim sqlcount, cnt, i
	sqlcount = "select count(jukyocd) as cnt"			'�˻� ���ð��� �ش��ϴ� �ε������� �����´�
	sqlcount = sqlcount & " from db_user.dbo.tbl_mileage_gubun"
	sqlcount = sqlcount & " where 1=1 "	

	if frectisusing <> "" then 
		sqlcount = sqlcount & "and isusing = '" & frectisusing & "'"
	end if			
	if frectseachjukyocd <> "" then 
		sqlcount = sqlcount & "and jukyocd = '" & frectseachjukyocd & "'"
	end if	
		
	rsget.open sqlcount,dbget,1
	'response.write sqlcount&"<br>"
	FTotalCount = rsget("cnt")				'�ѷ��ڵ� ���� �ε���ī��Ʈ�� �ְ�
	rsget.close	
	
	dim sql 
	sql = "select top "& FPageSize*FCurrpage &" jukyocd,jukyoname,isusing"
	sql = sql & " from db_user.dbo.tbl_mileage_gubun"
	sql = sql & " where 1=1 "
	
	if frectisusing <> "" then 
		sql = sql & "and isusing = '" & frectisusing & "'"
	end if			
	if frectseachjukyocd <> "" then 
		sql = sql & "and jukyocd = '" & frectseachjukyocd & "'"
	end if	

	sql = sql & " order by jukyocd desc"
	
	rsget.pagesize = FPageSize
	rsget.open sql,dbget,1
	'response.write sql&"<br>"
	
	FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
	FTotalPage = CInt(FTotalCount\FPageSize) + 1	
	
	redim flist(FResultCount)
	i = 0 
	
	if not rsget.eof then
		rsget.absolutepage = FCurrPage
		do until rsget.eof
			set flist(i) = new Cmileageoneitem
			
			flist(i).fjukyocd = rsget("jukyocd")
			flist(i).fjukyoname = rsget("jukyoname")
			flist(i).fisusing = rsget("isusing")
			
		rsget.movenext
		i = i + 1
		loop		
	end if
	rsget.close
	end sub
'##############################################################################################	
	public sub fmileage_add
	
	dim sql 
	sql = "select jukyocd,jukyoname,isusing"
	sql = sql & " from db_user.dbo.tbl_mileage_gubun"
	sql = sql & " where 1=1 and jukyocd = '" & frectjukyocd & "'"
	
	rsget.open sql,dbget,1
	
	FTotalCount = rsget.RecordCount
	redim flist(FTotalCount)
	i = 0 
	
	if not rsget.eof then
		do until rsget.eof
			set flist(i) = new Cmileageoneitem
			
			flist(i).fjukyocd = rsget("jukyocd")
			flist(i).fjukyoname = rsget("jukyoname")
			flist(i).fisusing = rsget("isusing")
			
		rsget.movenext
		i = i + 1
		loop		
	end if
	rsget.close
	end sub
'##############################################################################################		
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
%>