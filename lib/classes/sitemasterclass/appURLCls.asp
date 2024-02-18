<%
'#######################################################
'	History	: 2014-08-18 ����ȭ ����
'	Description : APP URL ����
'#######################################################

Class APPURLItem

	public Fidx
	Public Furltitle
	public Furlcontent
	public FisUsing
	public Fregdate
	public Furlhitcount
	Public Furldiv
	Public Furlcomplete
	Public Fcatecode
	Public Fqrsn

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class APPURL
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectidx
	Public FRectIsUsing
	Public FRectkeyWd
	Public FRecturldiv

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Sub getappurl()
		dim sqlStr,addSql, i

		'�߰� ����
		if FRectidx<>"" then
			addSql = addSql + " and idx = '" + FRectidx + "'" + vbcrlf
		end if
		if Not(FRectIsUsing="A" or FRectIsUsing="") then
			addSql = addSql + " and isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectkeyWd<>"" then
			addSql = addSql + " and urltitle like '%" + FRectkeyWd + "%'" + vbcrlf
		end If
		if FRecturldiv<>"" then
			addSql = addSql + " and urldiv  = '" + FRecturldiv + "'" + vbcrlf
		end if

		'�Ѽ� ����
		sqlStr = "select count(idx), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_AppUrlList " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		'Response.write sqlStr

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		'���� ����
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " idx, urltitle, urlcontent, isUsing, regdate, urlhitcount , urldiv , urlcomplete , catecode , qrsn" + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_AppUrlList " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by idx desc"

		'Response.write sqlstr

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new APPURLItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Furltitle		= db2html(rsget("urltitle"))
				FItemList(i).Furlcontent	= db2html(rsget("urlcontent"))
				FItemList(i).Furldiv		= rsget("urldiv")
				FItemList(i).FisUsing		= rsget("isUsing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Furlhitcount	= rsget("urlhitcount")
				FItemList(i).Furlcomplete	= rsget("urlcomplete")
				FItemList(i).Fcatecode		= rsget("catecode")
				FItemList(i).Fqrsn			= rsget("qrsn")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Sub DrawSelectBoxAppUrlDiv(boxname,stats)
%>
<select name='<%=boxname%>' class='select' onchange="chklink(this.value)">
	<option value="0">�����ϼ���</option>
	<option value="1" <% if stats = "1" then response.write " selected" %>>��ǰ��</option>
	<option value="2" <% if stats = "2" then response.write " selected" %>>�̺�Ʈ</option>
	<option value="3" <% if stats = "3" then response.write " selected" %>>�귣��</option>
	<option value="4" <% if stats = "4" then response.write " selected" %>>ī�װ�</option>
	<option value="8" <% if stats = "8" then response.write " selected" %>>�ܺ�URL</option>
	<option value="9" <% if stats = "9" then response.write " selected" %>>Today</option>
	<option value="10" <% if stats = "10" then response.write " selected" %>>����Ʈ</option>
	<option value="11" <% if stats = "11" then response.write " selected" %>>��ٱ���</option>
</select>
<%
end Sub
%>
