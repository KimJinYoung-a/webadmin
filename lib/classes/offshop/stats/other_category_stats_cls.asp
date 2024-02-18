<%
'###########################################################
' Description : 수작업 카테고리 통계 클래스
' Hieditor : 2011.11.16 한용민 생성
'###########################################################

Class cothercate_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fsellcnt
	public fsellsum
	public fsuplysum
	public fcdlcode_nm
	public fcdmcode_nm
	public fcatecdm
	public fitemid
	public fitemoptionname
	public fitemoption
	public fitemname
	public fmakerid
	public fsellprice
	public fcatecdl
end class

class cothercate_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public frectothercdl
	public frectshopid	
	public FRectStartDay
	public FRectEndDay
	public FRectmakerid
	public frectdatefg
	public frectcatecdm
	public frectothercheck
	public FRectInc3pl
	
	'/admin/offshop/stats/other_category_stats.asp
	public sub getother_category()
		dim sqlStr,i , sqlsearch

		if frectothercdl = "" then exit sub

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		
		if frectothercdl = "070" then
			sqlStr = sqlStr + " (case when i.catecdm in ('030','020') then '슈즈(패션슈즈,디자이너슈즈)'"
			sqlStr = sqlStr + " 	when i.catecdm in ('160','120','070') then '쥬얼리(쥬얼리,헤어악세사리,디자이너쥬얼리)'"
			sqlStr = sqlStr + " 	when i.catecdm in ('150','010','140','202','201','200') then '가방(지갑,패션가방,캐주얼가방,파우치,백팩,디자이너가방)'"
			sqlStr = sqlStr + " 	when i.catecdm in ('050') then '패션소품'"
			sqlStr = sqlStr + " 	when i.catecdm in ('040') then '디자인우산'"
			sqlStr = sqlStr + " 	when i.catecdm in ('110') then '패션시계'"
			sqlStr = sqlStr + " 	when i.catecdm in ('203') then '아이웨어'"
			sqlStr = sqlStr + " end) as cdmcode_nm"
			sqlStr = sqlStr + " ,(case when i.catecdm in ('030','020') then '030020'"
			sqlStr = sqlStr + " 	when i.catecdm in ('160','120','070') then '160120070'"
			sqlStr = sqlStr + " 	when i.catecdm in ('150','010','140','202','201','200') then '150010140202201200'"
			sqlStr = sqlStr + " 	when i.catecdm in ('050') then '050'"
			sqlStr = sqlStr + " 	when i.catecdm in ('040') then '040'"
			sqlStr = sqlStr + " 	when i.catecdm in ('110') then '110'"
			sqlStr = sqlStr + " 	when i.catecdm in ('203') then '203'"
			sqlStr = sqlStr + " end) as catecdm"
			sqlStr = sqlStr + " ,isnull(sum(case when i.catecdm in ('030','020') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('160','120','070') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('150','010','140','202','201','200') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('050') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('040') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('110') then d.itemno"
			sqlStr = sqlStr + " 	when i.catecdm in ('203') then d.itemno"
			sqlStr = sqlStr + " end),0) as sellcnt"
			sqlStr = sqlStr + " ,isnull(sum(case when i.catecdm in ('030','020') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('160','120','070') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('150','010','140','202','201','200') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('050') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('040') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('110') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('203') then d.itemno * d.realsellprice"
			sqlStr = sqlStr + " end),0) as sellsum"
			sqlStr = sqlStr + " ,isnull(sum(case when i.catecdm in ('030','020') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('160','120','070') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('150','010','140','202','201','200') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('050') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('040') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('110') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " 	when i.catecdm in ('203') then d.itemno * d.suplyprice"
			sqlStr = sqlStr + " end),0) as suplysum"
			sqlStr = sqlStr + " ,i.catecdl ,cl.code_nm as cdlcode_nm"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m "
			sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shopjumun_detail d "
			sqlStr = sqlStr + " 	on m.idx = d.masteridx"
			sqlStr = sqlStr + " 	and m.cancelyn='N'"
			sqlStr = sqlStr + " 	and d.cancelyn='N'"
			sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i "
			sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun "
			sqlStr = sqlStr + " 	and d.itemid=i.shopitemid "
			sqlStr = sqlStr + " 	and d.itemoption=i.itemoption "
			sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
		    sqlStr = sqlStr & "       on m.shopid=p.id "
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl"
			sqlStr = sqlStr + " 	on i.catecdl=cl.code_large "
			sqlStr = sqlStr + " where isnull(i.catecdl,'')<>'' and isnull(i.catecdm,'')<>'' and isnull(i.catecdn,'')<>''"
			sqlStr = sqlStr + " and i.catecdl='" + frectothercdl + "' " & sqlsearch
			sqlStr = sqlStr + " group by"
			sqlStr = sqlStr + " 	(case when i.catecdm in ('030','020') then '슈즈(패션슈즈,디자이너슈즈)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('160','120','070') then '쥬얼리(쥬얼리,헤어악세사리,디자이너쥬얼리)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('150','010','140','202','201','200') then '가방(지갑,패션가방,캐주얼가방,파우치,백팩,디자이너가방)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('050') then '패션소품'"
			sqlStr = sqlStr + " 		when i.catecdm in ('040') then '디자인우산'"
			sqlStr = sqlStr + " 		when i.catecdm in ('110') then '패션시계'"
			sqlStr = sqlStr + " 		when i.catecdm in ('203') then '아이웨어'"
			sqlStr = sqlStr + " 	end)"
			sqlStr = sqlStr + " 	,(case when i.catecdm in ('030','020') then '030020'"
			sqlStr = sqlStr + " 		when i.catecdm in ('160','120','070') then '160120070'"
			sqlStr = sqlStr + " 		when i.catecdm in ('150','010','140','202','201','200') then '150010140202201200'"
			sqlStr = sqlStr + " 		when i.catecdm in ('050') then '050'"
			sqlStr = sqlStr + " 		when i.catecdm in ('040') then '040'"
			sqlStr = sqlStr + " 		when i.catecdm in ('110') then '110'"
			sqlStr = sqlStr + " 		when i.catecdm in ('203') then '203'"
			sqlStr = sqlStr + " 	end)"
			sqlStr = sqlStr + " ,i.catecdl, cl.code_nm"
			sqlStr = sqlStr + " having "
			sqlStr = sqlStr + " 	(case when i.catecdm in ('030','020') then '슈즈(패션슈즈,디자이너슈즈)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('160','120','070') then '쥬얼리(쥬얼리,헤어악세사리,디자이너쥬얼리)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('150','010','140','202','201','200') then '가방(지갑,패션가방,캐주얼가방,파우치,백팩,디자이너가방)'"
			sqlStr = sqlStr + " 		when i.catecdm in ('050') then '패션소품'"
			sqlStr = sqlStr + " 		when i.catecdm in ('040') then '디자인우산'"
			sqlStr = sqlStr + " 		when i.catecdm in ('110') then '패션시계'"
			sqlStr = sqlStr + " 		when i.catecdm in ('203') then '아이웨어'"
			sqlStr = sqlStr + " 	end) is not null"
			sqlStr = sqlStr + " order by sellsum desc"
	
	elseif frectothercdl = "toms001" then
			sqlStr = sqlStr + " i.catecdl ,cl.code_nm as cdlcode_nm ,i.catecdm ,c2.code_nm as cdmcode_nm"
			sqlStr = sqlStr + " ,isnull(sum(d.itemno),0) as sellcnt"
			sqlStr = sqlStr + " ,isnull(sum(d.itemno * d.realsellprice),0) as sellsum"
			sqlStr = sqlStr + " ,isnull(sum(d.itemno * d.suplyprice),0) as suplysum"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m "
			sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shopjumun_detail d "
			sqlStr = sqlStr + " 	on m.idx = d.masteridx"
			sqlStr = sqlStr + " 	and m.cancelyn='N'"
			sqlStr = sqlStr + " 	and d.cancelyn='N'"
			sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i "
			sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun "
			sqlStr = sqlStr + " 	and d.itemid=i.shopitemid "
			sqlStr = sqlStr + " 	and d.itemoption=i.itemoption "
			sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
		    sqlStr = sqlStr & "       on m.shopid=p.id "			
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl"
			sqlStr = sqlStr + " 	on i.catecdl=cl.code_large "
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_mid c2"
			sqlStr = sqlStr + " 	on i.catecdl=c2.code_large "
			sqlStr = sqlStr + " 	and i.catecdm=c2.code_mid"
			sqlStr = sqlStr + " where isnull(i.catecdl,'')<>'' and isnull(i.catecdm,'')<>'' and isnull(i.catecdn,'')<>''"
			sqlStr = sqlStr + " and d.makerid='" + frectothercdl + "' " & sqlsearch
			sqlStr = sqlStr + " group by i.catecdl ,cl.code_nm ,i.catecdm ,c2.code_nm"
			sqlStr = sqlStr + " order by i.catecdl asc ,i.catecdm asc ,sellsum desc"			
		end if
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cothercate_oneitem
				
				FItemList(i).fcatecdl = rsget("catecdl")
				FItemList(i).fcatecdm = rsget("catecdm")
				FItemList(i).fsellcnt = rsget("sellcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplysum = rsget("suplysum")
				FItemList(i).fcdlcode_nm = db2html(rsget("cdlcode_nm"))
				FItemList(i).fcdmcode_nm = db2html(rsget("cdmcode_nm"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	'/admin/offshop/stats/other_category_stats_best.asp
	public sub getother_category_best()
		dim sqlStr,i , sqlsearch
		
		if frectothercdl = "" then exit sub

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"	
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if
		
		if frectothercheck = "ON" then
			if FRectmakerid="toms001" then
				if frectothercdl <> "" then
					sqlsearch = sqlsearch + " and i.catecdl='" + frectothercdl + "'"
				end if
				if frectcatecdm <> "" then
					sqlsearch = sqlsearch + " and i.catecdm='" + frectcatecdm + "'"
				end if
			end if
		else
			if frectothercdl="070" then
				sqlsearch = sqlsearch + " and i.catecdl='" + frectothercdl + "'"
	
				if frectcatecdm = "030020" then
					sqlsearch = sqlsearch + " and i.catecdm in ('030','020')"
				elseif frectcatecdm = "160120070" then
					sqlsearch = sqlsearch + " and i.catecdm in ('160','120','070')"
				elseif frectcatecdm = "150010140202201200" then
					sqlsearch = sqlsearch + " and i.catecdm in ('150','010','140','202','201','200')"
				elseif frectcatecdm = "050" then
					sqlsearch = sqlsearch + " and i.catecdm in ('050')"
				elseif frectcatecdm = "040" then
					sqlsearch = sqlsearch + " and i.catecdm in ('040')"
				elseif frectcatecdm = "110" then
					sqlsearch = sqlsearch + " and i.catecdm in ('110')"
				elseif frectcatecdm = "203" then
					sqlsearch = sqlsearch + " and i.catecdm in ('203')"			
				end if
			end if
		end if
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " d.itemid , d.itemoptionname , d.itemoption, d.itemname, d.makerid,d.sellprice"
		sqlStr = sqlStr + " ,sum(d.itemno) as sellcnt"
		sqlStr = sqlStr + " ,sum(d.itemno*d.realsellprice) as sellsum"
		sqlStr = sqlStr + " ,sum(d.itemno*d.suplyprice)as suplysum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m "
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shopjumun_detail d "
		sqlStr = sqlStr + " 	on m.idx = d.masteridx "
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid "
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "		
		sqlStr = sqlStr + " where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr + " and isnull(i.catecdl,'')<>'' and isnull(i.catecdm,'')<>'' and isnull(i.catecdn,'')<>''"
		sqlStr = sqlStr + " group by d.itemid, d.sellprice, d.itemname, d.makerid, d.itemoptionname, d.itemoption ,d.itemno"
		sqlStr = sqlStr + " order by sellsum desc, sellcnt Desc"		
			
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cothercate_oneitem
				
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = db2html(rsget("makerid"))
				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).fsellcnt = rsget("sellcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplysum = rsget("suplysum")

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
end Class

function other_category(boxname,selectedid,changeflg)
%>
	<select name="<%=boxname%>" <%=changeflg%>>
		<option value="" <% if selectedid="" then response.write " selected" %>>선택</option>
		<option value="070" <% if selectedid="070" then response.write " selected" %>>가방/슈즈/쥬얼리</option>
		<option value="toms001" <% if selectedid="toms001" then response.write " selected" %>>탐스슈즈전용</option>
	</select>
<%
end function
%>	