<%
Class Multi3Obj
'ÄÜÅÙÃ÷
public C_idx
public C_evt_code
public C_main_copy
public C_sub_copy
public C_main_color
public C_main_content
public C_regdate
public C_background_img
public C_reg_name
public C_content_order
public C_moddate
public C_mod_name

'ÄÜÅÙÃ÷À¯´Ö
public U_idx
public U_evt_code
public U_content_idx
public U_unit_class
public U_unit_order
public U_unit_main_copy
public U_unit_main_content
public U_tag
public U_regdate
public U_reg_name
public U_moddate
public U_mod_name
 
'¾ÆÀÌÅÛ
public I_idx
public I_evt_code
public I_itemid
public I_unit_idx
public I_item_img
public I_item_name
public I_item_order
public I_regdate
public I_reg_name
public I_moddate
public I_mod_name

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Multi3
    public FOneContent
    public FItemList()
	public FUnitList()	

	public FRectContentId
	public FRectUnitIdx
	public FRectEvtCode

	public FItemTotalCount
	public FUnitTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount          
	
	public FContentIdx
	
    public Sub getContentsUnitList()
        dim sqlStr, i, sqlWhere

		if FRectContentId <> "" then
			sqlwhere = " and content_idx = " & FRectContentId
		end if
		

		sqlStr = " select count(idx) as cnt from db_event.dbo.[tbl_multi3_content_units] "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere

        rsget.Open sqlStr, dbget, 1
			FUnitTotalCount = rsget("cnt")
		rsget.close
        
        if FUnitTotalCount < 1 then exit Sub        	
			
        sqlStr = "select * "
        sqlStr = sqlStr + " from db_event.dbo.[tbl_multi3_content_units] "
        sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere        
		sqlStr = sqlStr + " order by unit_order asc" 		

		redim preserve FUnitList(FUnitTotalCount)
		rsget.Open sqlStr, dbget, 1
		if not rsget.EOF  then
		    i = 0			
			do until rsget.eof
				set FUnitList(i) = new Multi3Obj

				FUnitList(i).U_idx				= rsget("idx")
				FUnitList(i).U_evt_code			= rsget("evt_code")
				FUnitList(i).U_content_idx		= rsget("content_idx")
				FUnitList(i).U_unit_class		= rsget("unit_class")
				FUnitList(i).U_unit_order		= rsget("unit_order")
				FUnitList(i).U_unit_main_copy	= rsget("unit_main_copy")
				FUnitList(i).U_unit_main_content= rsget("unit_main_content")
				FUnitList(i).U_tag				= rsget("tag")
				FUnitList(i).U_regdate			= rsget("regdate")
				FUnitList(i).U_reg_name			= rsget("reg_name")
				FUnitList(i).U_moddate			= rsget("moddate")
				FUnitList(i).U_mod_name			= rsget("mod_name")															

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub getUnitItemsList()
        dim sqlStr, i, sqlWhere

		if FRectUnitIdx <> "" then
			sqlwhere = " and unit_idx = " & FRectUnitIdx
		end if
		

		sqlStr = " select count(idx) as cnt from db_event.dbo.[tbl_multi3_items] "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere

        rsget.Open sqlStr, dbget, 1
			FItemTotalCount = rsget("cnt")
		rsget.close
        
        if FItemTotalCount < 1 then exit Sub        	
			
        sqlStr = "select * "
        sqlStr = sqlStr + " from db_event.dbo.[tbl_multi3_items] "
        sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere        
		sqlStr = sqlStr + " order by item_order asc" 		
		rsget.Open sqlStr, dbget, 1
		redim preserve FItemList(FItemTotalCount)

		if  not rsget.EOF  then
		    i = 0			
			do until rsget.eof
				set FItemList(i) = new Multi3Obj

				FItemList(i).I_idx			= rsget("idx")
				FItemList(i).I_evt_code		= rsget("evt_code")
				FItemList(i).I_itemid		= rsget("itemid")
				FItemList(i).I_unit_idx		= rsget("unit_idx")
				FItemList(i).I_item_img		= rsget("item_img")
				FItemList(i).I_item_name	= rsget("item_name")
				FItemList(i).I_item_order	= rsget("item_order")
				FItemList(i).I_regdate		= rsget("regdate")
				FItemList(i).I_reg_name		= rsget("reg_name")
				FItemList(i).I_moddate		= rsget("moddate")
				FItemList(i).I_mod_name		= rsget("mod_name")																	

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Sub GetOneContent()
		dim SqlStr
        sqlStr = " Select * "		
        sqlStr = sqlStr & " From db_event.dbo.[tbl_multi3_contents] "
        SqlStr = SqlStr & " where evt_code = " + FRectEvtCode

		'response.write sqlStr &"<br>"
		'response.end

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneContent = new Multi3Obj

        if Not rsget.Eof then
            FOneContent.C_idx				= rsget("idx")	
            FOneContent.C_evt_code			= rsget("evt_code")	
            FOneContent.C_main_copy			= rsget("main_copy")
			FOneContent.C_sub_copy			= rsget("sub_copy")			
            FOneContent.C_main_color		= rsget("main_color")	
            FOneContent.C_main_content		= rsget("main_content")				
            FOneContent.C_regdate			= rsget("regdate")				
			FOneContent.C_background_img	= rsget("background_img")				
			FOneContent.C_reg_name			= rsget("reg_name")				
			FOneContent.C_content_order		= rsget("content_order")			
			FOneContent.C_moddate			= rsget("moddate")			
			FOneContent.C_mod_name			= rsget("mod_name")			
        end if
        rsget.close
	End Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)
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
%>