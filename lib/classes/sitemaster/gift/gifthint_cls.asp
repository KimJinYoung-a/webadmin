<%
'###########################################################
' Description :  기프트
' History : 2015.01.26 한용민 생성
'###########################################################

Class Cgifthint_oneitem
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

    public fexecutetime
    public fisusing
    public forderno
    public fregdate
    public flastadminid
    public flastupdate
    public fitemidx
    public fitemid
    public fthemetype
    public ftitle
	public Fthemeidx
	public Fitemcnt
	public Fexecutedate
	public FsmallImage
	public fitemname
	public fitemscore
	public ftalkcount
end Class

Class Cgifthint
    public FOneItem
    public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public frectthemeidx
	public frectisusing
	public frecttitle
	public frectthemetype
	public frectitemid
	public frectexecutedate

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
    End Sub

	'/admin/sitemaster/gift/hint/gifthint_item.asp
    public Sub getgifthint_item()
       dim sqlStr, sqlsearch, i

		if frectthemeidx="" then exit Sub

		if frectitemid<>"" then
			sqlsearch = sqlsearch & " and s.itemid="& frectitemid &""
		end if
		if frectthemeidx<>"" then
			sqlsearch = sqlsearch & " and s.themeidx="& frectthemeidx &""
		end if
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and s.isusing='"& frectisusing &"'"
		end if
		if frectexecutedate<>"" then
			sqlsearch = sqlsearch & " and s.executedate='"& frectexecutedate &"'"
		end if

		sqlStr = " select count(s.itemid) as cnt"
		sqlStr = sqlStr & " from db_board.dbo.tbl_gifthint_item s"
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "		and i.itemid<>0 "
		sqlStr = sqlStr & "		and i.isusing='Y'"
		sqlStr = sqlStr & "		and i.sellyn in ('Y')"
		sqlStr = sqlStr & "		and i.sellcash>=10000"
		sqlStr = sqlStr & "		and (i.sellyn='Y' and (100-i.sellcash/i.orgprice*100) < 70)"        
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.itemidx, s.themeidx, s.itemid, s.executedate, s.isusing, s.orderno, s.regdate, s.lastadminid, s.lastupdate"
		sqlStr = sqlStr & " , i.itemname, i.smallImage, i.itemscore"
		sqlStr = sqlStr & " ,(select count(talk_idx) as talkcount from db_board.dbo.tbl_shopping_talk_item where s.itemid=itemid) as talkcount"
		sqlStr = sqlStr & " from db_board.dbo.tbl_gifthint_item s"
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "		and i.itemid<>0 "
		sqlStr = sqlStr & "		and i.isusing='Y'"
		sqlStr = sqlStr & "		and i.sellyn in ('Y')"
		sqlStr = sqlStr & "		and i.sellcash>=10000"
		sqlStr = sqlStr & "		and (i.sellyn='Y' and (100-i.sellcash/i.orgprice*100) < 70)"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by (case when s.orderno<>99 then 0 else 1 end) asc, i.itemscore desc"

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cgifthint_oneitem
				
				FItemList(i).ftalkcount = rsget("talkcount")
				FItemList(i).fitemscore = rsget("itemscore")
				FItemList(i).Fitemidx = rsget("itemidx")
				FItemList(i).Fthemeidx = rsget("themeidx")
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fexecutedate = rsget("executedate")
				FItemList(i).Fisusing = rsget("isusing")
				FItemList(i).Forderno = rsget("orderno")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastadminid = rsget("lastadminid")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).fitemname = db2html(rsget("itemname"))
	            FItemList(i).FsmallImage = chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//admin/sitemaster/gift/hint/gifthint.asp
    public Sub getgifthint_one()
        dim SqlStr, sqlsearch

		if frectthemeidx="" then exit Sub

		if frectthemeidx<>"" then
			sqlsearch = sqlsearch & " and t.themeidx="& frectthemeidx &""
		end if
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and t.isusing='"& frectisusing &"'"
		end if

        SqlStr = "select top 1"
        sqlstr = sqlstr & " t.themeidx, t.themetype, t.title, t.executetime, t.isusing, t.orderno, t.regdate, t.lastadminid, t.lastupdate"
        sqlstr = sqlstr & " from db_board.dbo.tbl_gifthint t with (nolock)"
        sqlstr = sqlstr & " where 1=1 " & sqlsearch
		
		'response.write sqlstr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget.RecordCount

        set FOneItem = new Cgifthint_oneitem
        if Not rsget.Eof then

            FOneItem.Fthemeidx = rsget("themeidx")
            FOneItem.Fthemetype = rsget("themetype")
            FOneItem.fexecutetime = rsget("executetime")
            FOneItem.Ftitle = db2html(rsget("title"))
            FOneItem.Fisusing = rsget("isusing")
            FOneItem.Forderno = rsget("orderno")
            FOneItem.Fregdate = rsget("regdate")
            FOneItem.Flastadminid = rsget("lastadminid")
            FOneItem.Flastupdate = rsget("lastupdate")

        end if
        rsget.close
    end Sub

	'//admin/sitemaster/gift/hint/gifthint.asp
    public Sub getgifthint_list()
        dim sqlStr, sqlsearch

		if frectthemeidx<>"" then
			sqlsearch = sqlsearch & " and t.themeidx="& frectthemeidx &""
		end if
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and t.isusing='"& frectisusing &"'"
		end if
		if frecttitle<>"" then
			sqlsearch = sqlsearch & " and t.title like '%"& frecttitle &"%'"
		end if
		if frectthemetype<>"" then
			sqlsearch = sqlsearch & " and t.themetype = "& frectthemetype &""
		end if

        sqlStr = "select count(t.themeidx) as cnt"
        sqlstr = sqlstr & " from db_board.dbo.tbl_gifthint t"
        sqlStr = sqlStr & " Where 1=1 " & sqlsearch

		'response.write sqlstr & "<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
		
		if FTotalCount < 1 then exit Sub

        sqlStr = "select top " + CStr(FPageSize * FCurrPage)
        sqlstr = sqlstr & " t.themeidx, t.themetype, t.title, t.executetime, t.isusing, t.orderno, t.regdate, t.lastadminid, t.lastupdate"
        sqlstr = sqlstr & " from db_board.dbo.tbl_gifthint t"
        sqlStr = sqlStr & " Where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by t.executetime desc"

		'response.write sqlstr & "<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cgifthint_oneitem
				
	            FItemList(i).Fthemeidx = rsget("themeidx")
	            FItemList(i).Fthemetype = rsget("themetype")
	            FItemList(i).fexecutetime = rsget("executetime")
	            FItemList(i).Ftitle = db2html(rsget("title"))
	            FItemList(i).Fisusing = rsget("isusing")
	            FItemList(i).Forderno = rsget("orderno")
	            FItemList(i).Fregdate = rsget("regdate")
	            FItemList(i).Flastadminid = rsget("lastadminid")
	            FItemList(i).Flastupdate = rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

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

function drawthemetype(boxname, themetype, chval)
%>
	<select name="<%= boxname %>" <%= chval %>>
		<option value="" <% if themetype="" then response.write " selected" %>>선택</option>
		<option value="1" <% if themetype="1" then response.write " selected" %>>HIM</option>
		<option value="2" <% if themetype="2" then response.write " selected" %>>TEEN</option>
		<option value="3" <% if themetype="3" then response.write " selected" %>>BABY</option>
		<option value="4" <% if themetype="4" then response.write " selected" %>>HER</option>
		<option value="5" <% if themetype="5" then response.write " selected" %>>HOME</option>
	</select>
<%
end function

function getthemetype(themetype)
	dim tmpthemetype

	if themetype="" then exit function

	if themetype="1" then
		tmpthemetype="HIM"
	elseif themetype="2" then
		tmpthemetype="TEEN"
	elseif themetype="3" then
		tmpthemetype="BABY"
	elseif themetype="4" then
		tmpthemetype="HER"
	elseif themetype="5" then
		tmpthemetype="HOME"
	end if
	
	getthemetype=tmpthemetype
end function
%>