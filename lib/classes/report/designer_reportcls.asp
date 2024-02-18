<%
class CDesignerJumunList

 	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
	public FCancelyn
    public FItemCount
	public FItemID
	public FItemName
	public FRectDesignerID
	public FRectFromDate
	public FRectToDate
	public FRectSettle2
	public Fitemimage
	public FRectOldJumun
	public FRectItemID
	public FRectItemName
	
	Private Sub Class_Initialize()
	redim FMasterItemList(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public sub SearchItemPort()

        Dim sql, sqltmp, wheredetail, i
    
        maxt = -1
        maxc = -1

		if FRectsettle2="m" then
		    ''월별
			sqltmp = " d.itemid,d.itemname, i.smallimage,convert(varchar(7),m.regdate,21) "
		else
		    ''일별
			sqltmp = " d.itemid,d.itemname, i.smallimage,convert(varchar(10),m.regdate,21) "
		end if

		
		''#################################################
		''데이타.
		''#################################################


			sql = "select top 1000 " & sqltmp & " as ipdate,  sum(d.buycash*d.itemno) as sumtotal,"
			sql = sql + " sum(d.itemno) as sellcnt from "

			if FRectOldJumun="on" then
				sql = sql + "  [db_log].[dbo].tbl_old_order_master_2003 m "
				sql = sql + "       Join [db_log].[dbo].tbl_old_order_detail_2003 d"
				sql = sql + "       on m.orderserial = d.orderserial"
				if (FRectDesignerID<>"") then
        			sql = sql + " and d.makerid='" + FRectDesignerID + "'"
        		end if
    		
			else
				sql = sql + "  [db_order].[dbo].tbl_order_master m "
				sql = sql + "       Join [db_order].[dbo].tbl_order_detail d"
				sql = sql + "       on m.orderserial = d.orderserial"
				if (FRectDesignerID<>"") then
        			sql = sql + " and d.makerid='" + FRectDesignerID + "'"
        		end if
			end if
			sql = sql + "   left join [db_item].[dbo].tbl_item i"
			sql = sql + "   on d.itemid = i.itemid "
			sql = sql + " where (m.regdate >= '" & FRectFromDate & "') and (m.regdate <='" & FRectToDate & "')"
			
    		if (FRectItemID<>"") then
    			sql = sql + " and d.itemid= " & FRectItemID  
    		end if
			if (FRectItemName<>"") then
				sql = sql + " and d.itemname like '%" & FRectItemName & "%'"
			end if
			
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and d.cancelyn<>'Y'"
            sql = sql + " Group by " + sqltmp
	        sql = sql & " Order by d.itemid, ipdate  "
            
            if FRectOldJumun="on" then ''과거주문 데이타 마트 디비로 //2014/12/02
                db3_rsget.CursorLocation = adUseClient
    			db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
    			
    			FResultCount = db3_rsget.RecordCount
	            redim preserve FMasterItemList(FResultCount)

				do until db3_rsget.eof

					set FMasterItemList(i) = new CDesignerJumunList

					FMasterItemList(i).Fitemimage = db3_rsget("smallimage")
					FMasterItemList(i).Fitemid = db3_rsget("itemid")
					FMasterItemList(i).Fitemname = db2html(db3_rsget("itemname"))
					FMasterItemList(i).Fseldate = db3_rsget("ipdate")
					FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
					FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

					if Not IsNull(FMasterItemList(i).Fselltotal) then
						maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
						maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					end if

					db3_rsget.MoveNext
					i = i + 1
				loop

		        db3_rsget.close
            else
                rsget.CursorLocation = adUseClient
    			rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    			
    			FResultCount = rsget.RecordCount
	            redim preserve FMasterItemList(FResultCount)

				do until rsget.eof

					set FMasterItemList(i) = new CDesignerJumunList

					FMasterItemList(i).Fitemimage = rsget("smallimage")
					FMasterItemList(i).Fitemid = rsget("itemid")
					FMasterItemList(i).Fitemname = db2html(rsget("itemname"))
					FMasterItemList(i).Fseldate = rsget("ipdate")
					FMasterItemList(i).Fselltotal = rsget("sumtotal")
					FMasterItemList(i).Fsellcnt = rsget("sellcnt")

					if Not IsNull(FMasterItemList(i).Fselltotal) then
						maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
						maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					end if

					rsget.MoveNext
					i = i + 1
				loop

		        rsget.close
    	    end if
	end sub

end class



class CJumunMasterItem
	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
    public FItemCount
	public FItemID
	public FItemName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub


end Class

class CJumunMaster
	public FMasterItemList()
	public maxt
	public maxc
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
    public FRectItemid
	public FCurrPage
	public FRectSettle
	public FRectSettle2
	public FRectFromDate
	public FRectToDate
	public FRectIpkumDiv4
    public FRectDesignerID
    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal
	public FRectOldJumun

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public sub SearchSellrePort()

        Dim sql, sqltmp, wheredetail, i

	    maxt = -1
	    maxc = -1

		if FRectsettle2="m" then
			sqltmp = " convert(varchar(7),m.regdate,120)"
		else
			sqltmp = " convert(varchar(10),m.regdate,120) "
		end if

		

		''#################################################
		''데이타.
		''#################################################


			sql = "select top 1000 " & sqltmp & " as ipdate,  sum(d.buycash*d.itemno) as sumtotal,"
			sql = sql + " count(m.orderserial) as sellcnt from "

			if FRectOldJumun="on" then
				sql = sql + "  [db_log].[dbo].tbl_old_order_master_2003 m "
				sql = sql + "       Join [db_log].[dbo].tbl_old_order_detail_2003 d"
				sql = sql + "       on m.orderserial = d.orderserial"
				
			else
				sql = sql + "  [db_order].[dbo].tbl_order_master m "
				sql = sql + "       Join [db_order].[dbo].tbl_order_detail d"
				sql = sql + "       on m.orderserial = d.orderserial"
				
			end if
			sql = sql + " where (m.regdate >= '" & FRectFromDate & "') and (m.regdate <='" & FRectToDate & "')"
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and d.cancelyn<>'Y'"
			sql = sql + " and d.itemid <> 0"
            if (FRectDesignerID<>"") then
    		    sql = sql + " and d.makerid='" + FRectDesignerID + "'"
    		end if
    		
            sql = sql + " Group by " + sqltmp
	        sql = sql & " Order by ipdate "

            if FRectOldJumun="on" then ''과거주문 데이타 마트 디비로 //2014/12/02
                db3_rsget.CursorLocation = adUseClient
                db3_dbget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
    			db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
    
                FResultCount = db3_rsget.RecordCount
    
    	        redim preserve FMasterItemList(FResultCount)
    			do until db3_rsget.eof
    
    		        set FMasterItemList(i) = new CDesignerJumunList
    				FMasterItemList(i).Fseldate = db3_rsget("ipdate")
    				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
    				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
    
    				if Not IsNull(FMasterItemList(i).Fselltotal) then
    					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
    					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
    				end if
    
    				db3_rsget.MoveNext
    				i = i + 1
    			loop
    		    db3_rsget.close
            else
                rsget.CursorLocation = adUseClient
    			rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    
                FResultCount = rsget.RecordCount
    
    	        redim preserve FMasterItemList(FResultCount)
    			do until rsget.eof
    
    		        set FMasterItemList(i) = new CDesignerJumunList
    				FMasterItemList(i).Fseldate = rsget("ipdate")
    				FMasterItemList(i).Fselltotal = rsget("sumtotal")
    				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
    
    				if Not IsNull(FMasterItemList(i).Fselltotal) then
    					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
    					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
    				end if
    
    				rsget.MoveNext
    				i = i + 1
    			loop
    		    rsget.close
		    end if
		    
		if maxt=0 then maxt=1
		if maxc=0 then maxc=1
	end sub
    
    public sub SearchSellrePort_HPDCase()

        Dim sql, sqltmp, wheredetail, i

	    maxt = -1
	    maxc = -1

		if FRectsettle2="m" then
			sqltmp = " convert(varchar(7),m.regdate,120)"
		else
			sqltmp = " convert(varchar(10),m.regdate,120) "
		end if

		

		''#################################################
		''데이타.
		''#################################################


			sql = "select top 1000 " & sqltmp & " as ipdate,  sum(d.itemcost*d.itemno) as sumtotal,"
			sql = sql + " count(m.orderserial) as sellcnt from "

			if FRectOldJumun="on" then
				sql = sql + "  [db_log].[dbo].tbl_old_order_master_2003 m "
				sql = sql + "       Join [db_log].[dbo].tbl_old_order_detail_2003 d"
				sql = sql + "       on m.orderserial = d.orderserial"
				
			else
				sql = sql + "  [db_order].[dbo].tbl_order_master m "
				sql = sql + "       Join [db_order].[dbo].tbl_order_detail d"
				sql = sql + "       on m.orderserial = d.orderserial"
				
			end if
			sql = sql + " where (m.regdate >= '" & FRectFromDate & "') and (m.regdate <='" & FRectToDate & "')"
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and d.cancelyn<>'Y'"
			sql = sql + " and d.itemid <> 0"
    		sql = sql + " and d.makerid='" + FRectDesignerID + "'"
    		
            sql = sql + " Group by " + sqltmp
	        sql = sql & " Order by ipdate "

            rsget.CursorLocation = adUseClient
			rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

			FResultCount = rsget.RecordCount

		    redim preserve FMasterItemList(FResultCount)

			do until rsget.eof

				set FMasterItemList(i) = new CDesignerJumunList
					FMasterItemList(i).Fseldate = rsget("ipdate")
					FMasterItemList(i).Fselltotal = rsget("sumtotal")
					FMasterItemList(i).Fsellcnt = rsget("sellcnt")

					if Not IsNull(FMasterItemList(i).Fselltotal) then
						maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
						maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					end if

					rsget.MoveNext
					i = i + 1
				loop
			rsget.close
		if maxt=0 then maxt=1
		if maxc=0 then maxc=1
	end sub

	public Sub SearchJumunListByfavitemlist()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""


		if FRectsettle = "self" then
		wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectItemid<>"") then
		wheredetail = wheredetail + " and i.itemid=" + FRectItemid
		end if


		''#################################################
		''총 갯수. 총금액
		''#################################################
		sqlStr = "select DISTINCT count(m.itemid) as Titemid"
		sqlStr = sqlStr + " from  [db_my10x10].[dbo].tbl_myfavorite m, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where m.itemid = i.itemid"
		sqlStr = sqlStr + wheredetail


		rsget.Open sqlStr,dbget,1
		FSubtotal = rsget("Titemid")
		rsget.Close


		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select "
		sqlStr = sqlStr + "DISTINCT m.itemid, count(m.itemid) as itemcount, i.itemname, i.smallimage "
		sqlStr = sqlStr + " from [db_my10x10].[dbo].tbl_myfavorite m, [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " where m.itemid = i.itemid"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by m.itemid, i.itemname, i.smallimage"
		sqlStr = sqlStr + " order by itemcount desc"

'response.write sqlStr

		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
'response.write FPageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster
	'			FMasterItemList(i).Forderserial = rsget("orderserial")
	'			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
	'			FMasterItemList(i).Fuserid		= rsget("userid")
	'			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
	'			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
	'			FMasterItemList(i).Fregdate		= rsget("regdate")
	'			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
	'			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
	'			FMasterItemList(i).Freqphone	= rsget("reqphone")
	'			FMasterItemList(i).Freqhp		= rsget("reqhp")
	'			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
	'			FMasterItemList(i).Fsitename	= rsget("sitename")
	'			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
	'			FMasterItemList(i).FCancelyn	= rsget("cancelyn")
				FMasterItemList(i).FItemID       = rsget("itemid")
			    FMasterItemList(i).FItemCount       = rsget("itemcount")
    			FMasterItemList(i).FItemimgsmall      = rsget("smallimage")
				FMasterItemList(i).FItemName     = rsget("itemname")
	'			FMasterItemList(i).FItemOption   = rsget("itemoption")


						if Not IsNull(FMasterItemList(i).FItemCount) then
							maxc = MaxVal(maxc,FMasterItemList(i).FItemCount)
						end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub SearchJumunListBybestfavitemlist()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""


		if FRectsettle = "self" then
		wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectItemid<>"") then
		wheredetail = wheredetail + " and i.itemid=" + FRectItemid
		end if


		''#################################################
		''총 갯수. 총금액
		''#################################################

		sqlStr = "select sum(c.favcount) as favtotalcount"
		sqlStr = sqlStr + " from [db_const].[dbo].tbl_const_category c, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where c.itemid = i.itemid"
		sqlStr = sqlStr + wheredetail
		rsget.Open sqlStr,dbget,1

		FSubtotal = rsget("favtotalcount")


		rsget.Close



		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select "
		sqlStr = sqlStr + "c.itemid, c.favcount, i.itemname, g.imgsmall"
		sqlStr = sqlStr + " from [db_const].[dbo].tbl_const_category c, [db_item].[dbo].tbl_item i, [db_item].[dbo].tbl_item_image g"
		sqlStr = sqlStr + " where c.itemid = i.itemid"
		sqlStr = sqlStr + " and c.itemid = g.itemid"
		sqlStr = sqlStr + wheredetail
'		sqlStr = sqlStr + " group by c.itemid, c.favcount, i.itemname, g.imgsmall"
		sqlStr = sqlStr + " order by c.favcount desc"

'response.write sqlStr

		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster
	'			FMasterItemList(i).Forderserial = rsget("orderserial")
	'			FMasterItemList(i).Fjumundiv	= rsget("jumundiv")
	'			FMasterItemList(i).Fuserid		= rsget("userid")
	'			FMasterItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
	'			FMasterItemList(i).Fipkumdiv	= rsget("ipkumdiv")
	'			FMasterItemList(i).Fregdate		= rsget("regdate")
	'			FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
	'			FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
	'			FMasterItemList(i).Freqphone	= rsget("reqphone")
	'			FMasterItemList(i).Freqhp		= rsget("reqhp")
	'			FMasterItemList(i).Fdeliverno	= rsget("deliverno")
	'			FMasterItemList(i).Fsitename	= rsget("sitename")
	'			FMasterItemList(i).Fdiscountrate	= rsget("discountrate")
	'			FMasterItemList(i).FCancelyn	= rsget("cancelyn")
				FMasterItemList(i).FItemID       = rsget("itemid")
			    FMasterItemList(i).FItemCount       = rsget("favcount")
    			FMasterItemList(i).FItemimgsmall      = rsget("imgsmall")
				FMasterItemList(i).FItemName     = rsget("itemname")
	'			FMasterItemList(i).FItemOption   = rsget("itemoption")

						if Not IsNull(FMasterItemList(i).FItemCount) then
							maxc = MaxVal(maxc,FMasterItemList(i).FItemCount)
						end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


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
