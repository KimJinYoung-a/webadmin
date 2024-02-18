<%
Class CTSKeywordlItem

	public Fidx
	public Fitemid
	public Fcate_large
	public Fcate_mid
	Public FsortNo
	public Fregdate
	Public Ftitle
	Public FisUsing
	public Fcomment
	public Fuserid
	public Fitemname
	public Fsellcash

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CSearchKeyWord
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectIdx
	public FRectUsing
	public FRectSearch

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetSearchreview()
		dim sqlStr, addSql, i

		if FRectUsing<>"" then
			addSql = " Where a.isusing='" & FRectUsing & "'"
		else
			addSql = " Where a.isusing='Y'"
		end if

		if FRectSearch<>"" then
			addSql = addSql & " and a.itemId like '%" & FRectSearch & "%'"
		end if

		if FRectIdx<>"" then
			addSql = addSql & " and a.idx='" & FRectIdx & "'"
		end if

		sqlStr = "select count(a.idx) as cnt from [db_sitemaster].[dbo].tbl_main_review as a"
		sqlStr = sqlStr + addSql

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " a.idx,a.itemid,a.sortNo,a.isUsing,a.regdate, a.comment, a.userid , a.cate_mid, b.itemname, b.sellcash, " + vbcrlf
		sqlStr = sqlStr + " case a.cate_large when '10' then '�����ι���' when '20' then '���ǽ�/���μ�ǰ' when '25' then '������' when '30' then 'Ű��Ʈ'" + vbcrlf
		sqlStr = sqlStr + " when '35' then '����/���' when '40' then '����' when '45' then '����/��Ȱ' when '50' then 'Ȩ/����' when '55' then '�к긯' when '60' then 'Űģ' when '70' then '����/����/���'" + vbcrlf
		sqlStr = sqlStr + " when '75' then '��Ƽ' when '80' then 'Women' when '90' then 'Men' when '100' then '���̺�' when '110' then '����ä��' end as cate_large" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_review as a" + vbcrlf
		sqlStr = sqlStr + " join db_item.dbo.tbl_item as b" + vbcrlf
		sqlStr = sqlStr + " on a.itemid=b.itemid" + addSql + vbcrlf
		sqlStr = sqlStr + " order by sortNo asc, idx desc"
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTSKeywordlItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).FsortNo	= rsget("sortNo")
				FItemList(i).FisUsing	= rsget("isUsing")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Fcomment	= rsget("comment")
				FItemList(i).Fuserid	= rsget("userid")
				FItemList(i).Fcate_large= rsget("cate_large")
				FItemList(i).Fcate_mid	= rsget("cate_mid")
				FItemList(i).Fitemname	= rsget("itemname")
				FItemList(i).Fsellcash	= rsget("sellcash")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
