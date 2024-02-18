<%
'#############################################################
'	Description : Ŭ����� ���� ���� Ŭ����
'	History		: 2016.01.14 ���¿� ����
'#############################################################
%>
<%
class CHitchhikerItem
	public Fidx
	public FReqIsusing
	public FReqItemid
	public FreqRegdate
end class

class CClaearanceitem
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FIsusing
	public Fitemid

	'###### Ŭ��������� ��ǰ�ڵ� ����Ʈ ######
	public sub fnGetclaearanceitemList
		dim sqlStr,i, sqlsearch

		if Fitemid <> "" Then
			sqlsearch = sqlsearch & " AND itemid = '"& Fitemid &"'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if

		'���� �� ���� ���ϱ�
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_clearance_sale_item"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB ������ ����Ʈ
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, itemid, isusing, regdate"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_clearance_sale_item"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize		
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CHitchhikerItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FReqitemid = rsget("itemid")
					FItemList(i).FReqIsusing = rsget("isusing")
					FItemList(i).FreqRegdate = rsget("regdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class
%>






	

		