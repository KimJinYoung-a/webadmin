<%
'###############################################
' PageName :topcatecodeCls
' Discription : 모바일 GNB 카테고리 클레스
' History : 2015-09-14 이종화 생성
'###############################################

'// GNBCODE
Class GNBcodeitem
    Public Fgnbcode
    public Fgnbname
	Public Fdispcode '//카테고리 코드
	Public Fdispname '//카테고리 NAME
    Public Fisusing
	Public Fsortnum 
	Public Fidx
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
end Class 

Class GNBcode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectgnbcode
    
    public Sub GetOneContentsCode()
        dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from db_sitemaster.[dbo].[tbl_mobile_main_topcatecode] "
        SqlStr = SqlStr + " where gnbcode=" + CStr(FRectgnbcode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new GNBcodeitem
        if Not rsget.Eof then
            
            FOneItem.Fgnbcode		= rsget("gnbcode")
            FOneItem.Fgnbname	= db2html(rsget("gnbname"))
            FOneItem.Fisusing		= rsget("isusing")            

        end if
        rsget.close
    end Sub
    
    public Sub GetgnbcodeList()
        dim sqlStr
        sqlStr = "select count(gnbcode) as cnt from db_sitemaster.[dbo].[tbl_mobile_main_topcatecode]"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from db_sitemaster.[dbo].[tbl_mobile_main_topcatecode] "
        sqlStr = sqlStr + " order by gnbcode desc"
        
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
				set FItemList(i) = new GNBcodeitem

				FItemList(i).Fgnbcode		= rsget("gnbcode")
                FItemList(i).Fgnbname		= db2html(rsget("gnbname"))
                FItemList(i).Fisusing			= rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount     = 0
		FScrollCount      = 10
		FTotalCount        = 0

	End Sub

	Private Sub Class_Terminate()

    End Sub

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

'// GNB SUBCODE

Class GNBsubcode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public Fidx
	Public FRectgnbcode
	Public Fisusing
    
    public Sub GetOneSubCode()
        dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from db_sitemaster.[dbo].[tbl_mobile_main_topsubcode] "
        SqlStr = SqlStr + " where idx=" + CStr(Fidx) 

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new GNBcodeitem
        if Not rsget.Eof then
            
            FOneItem.Fgnbcode		= rsget("gnbcode")
            FOneItem.Fdispcode	= rsget("dispcode")
            FOneItem.Fisusing		= rsget("isusing")            

        end if
        rsget.close
    end Sub
    
    public Sub GetSubCodeList()
        dim sqlStr
        sqlStr = "select count(*) as cnt from db_sitemaster.[dbo].[tbl_mobile_main_topsubcode] as s "
		sqlStr = sqlStr + " inner join db_sitemaster.[dbo].[tbl_mobile_main_topcatecode] as t on s.gnbcode = t.gnbcode  and t.isusing = 'Y' "
		sqlStr = sqlStr + " inner join [db_item].[dbo].[tbl_display_cate] as d on s.dispcode = d.catecode and d.depth = 1 and d.useyn = 'Y' "
		sqlStr = sqlStr + " where 1=1"

		if Fisusing<>"" then
            sqlStr = sqlStr + " and s.isusing='" + CStr(Fisusing) + "'"
        end If

		If FRectgnbcode <> "" Then '//gnbcode
			sqlStr = sqlStr + " and s.gnbcode='" + CStr(FRectgnbcode) + "'"
		End If 
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * , s.idx from db_sitemaster.[dbo].[tbl_mobile_main_topsubcode] as s "
		sqlStr = sqlStr + " inner join db_sitemaster.[dbo].[tbl_mobile_main_topcatecode] as t on s.gnbcode = t.gnbcode  and t.isusing = 'Y' "
		sqlStr = sqlStr + " inner join [db_item].[dbo].[tbl_display_cate] as d on s.dispcode = d.catecode and d.depth = 1 and d.useyn = 'Y' "
		sqlStr = sqlStr + " where 1=1"

		if Fisusing<>"" then
            sqlStr = sqlStr + " and s.isusing='" + CStr(Fisusing) + "'"
        end If

		If FRectgnbcode <> "" Then '//gnbcode
			sqlStr = sqlStr + " and s.gnbcode='" + CStr(FRectgnbcode) + "'"
		End If 

        sqlStr = sqlStr + " order by s.idx desc"

'		Response.write sqlStr
'		Response.end
        
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
				set FItemList(i) = new GNBcodeitem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fgnbcode		= rsget("gnbcode")
				FItemList(i).Fgnbname		= db2html(rsget("gnbname"))
                FItemList(i).Fdispcode		= rsget("dispcode")
                FItemList(i).Fdispname		= db2html(rsget("catename"))
                FItemList(i).Fisusing			= rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount     = 0
		FScrollCount      = 10
		FTotalCount        = 0

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
end Class

'//gnbcate
Sub drawSelectBoxGNB(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select gnbcode, gnbname from db_sitemaster.[dbo].[tbl_mobile_main_topcatecode] "
   query1 = query1 + " where isusing='Y'"
   query1 = query1 + " order by gnbcode asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("gnbcode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" & rsget("gnbcode") & "' " & tmp_str & ">" & rsget("gnbcode") & " [" & db2html(rsget("gnbname")) & "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

'//dispcode
Sub drawSelectBoxDISP(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select catecode, catename from [db_item].[dbo].[tbl_display_cate] "
   query1 = query1 + " where depth = '1'  and useyn = 'Y' "
   query1 = query1 + " order by sortNo, catecode asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("catecode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" & rsget("catecode") & "' " & tmp_str & ">" & rsget("catecode") & " [" & db2html(rsget("catename")) & "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>