<%

class Cbaesong_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fitemid			'상품code
	public fitemoption		'옵션code
	public fdelayday		'배송소요일
	public fitemname		'상품명
	public foptionname		'옵션명
	public fyyyymmdd
	public fmakerID
	public fcdL
	public fcdM
	public fcdS
	public fsaleNo
	public fsaleCost
	public forderCnt
	public forderManCnt
	public forderWomanCnt
	public forderManAge
	public forderWomanAge
	public fpassday
	public fcatename
	public fyyyy
	public fmm

end class


class Cbaesong_list

	public FDBName
	
	Private Sub Class_Terminate()
	End Sub
	
	Private Sub Class_Initialize()
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
		IF application("Svr_Info") = "Dev" THEN
			FDBName		= "db_back"
		Else
			FDBName		= "db_datamart"
		End If
	End Sub
	
	
	public flist()

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FSDate
	public FEDate
	public FItemID
	public FMakerID
	public FCateLarge
	public FItemname
	public FIsNotZero

	public FGubun
	public FTop
	
	

	public function fbaesong_list			'상품 및 브랜드 일별 세부
	dim i , sql, strSubSql
	
		If FSDate <> "" AND FEDate <> "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd Between '" & FSDate & "' AND '" & FEDate & "' "
		ElseIf FSDate <> "" AND FEDate = "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd <= '" & FSDate & "' "
		ElseIf FSDate = "" AND FEDate <> "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd >= '" & FEDate & "' "
		End If
		
		If FItemID <> "" Then
			strSubSql = strSubSql & " AND m.itemID = '" & FItemID & "' "
		End If
		
		If FMakerID <> "" Then
			strSubSql = strSubSql & " AND m.makerID = '" & FMakerID & "' "
		End If
		
		If FCateLarge <> "" Then
			strSubSql = strSubSql & " AND m.cdL = '" & FCateLarge & "' "
		End If
		
		If FItemname <> "" Then
			strSubSql = strSubSql & " AND m.itemName Like '%" & FItemname & "%' "
		End If
		
		If FIsNotZero = "Y" Then
			strSubSql = strSubSql & " AND m.passday > '0' "
		End If

		sql = " SELECT COUNT(*) From " & _
			  " 		[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
			  "	WHERE " & _
			  "		1=1 " & strSubSql & " "
		db3_rsget.open sql,db3_dbget,1
		'response.write strSql
		IF not db3_rsget.EOF THEN
			FTotalCount = db3_rsget(0)
		END IF
		db3_rsget.close
		
		
		IF FTotalCount > 0 THEN
			sql = "SELECT TOP "& (FPageSize * FCurrPage) &" " & _
				  "			m.yyyymmdd, m.makerID, m.itemID, m.itemOption, m.itemName, m.optionName, " & _
				  "			m.cdLName, " & _
				  "			m.cdL, m.cdM, m.cdS, m.saleNo, m.saleCost, m.orderCnt,  " & _
				  "			m.orderManCnt, m.orderWomanCnt, m.orderManAge, m.orderWomanAge, m.passday " & _
				  "		FROM " & _
				  "			[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
				  "		WHERE " & _
				  "			1=1 " & strSubSql & " " & _
				  "		ORDER BY yyyymmdd DESC " & _
				  ""
			
			'response.write sql&"<br>"	
			db3_rsget.open sql,db3_dbget,1
			
		
			if (FCurrPage * FPageSize < FTotalCount) then
				FResultCount = FPageSize
			else
				FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
			end if
			
			FTotalPage = (FTotalCount\FPageSize)
			
			if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
			
			redim preserve flist(FResultCount)

			db3_rsget.PageSize= FPageSize
			If  not db3_rsget.EOF  then
				db3_rsget.absolutepage = FCurrPage
				Do Until db3_rsget.Eof
					set flist(i) = new Cbaesong_oneitem
						
						flist(i).fitemid 		= db3_rsget("itemid")
						flist(i).fitemoption 	= db3_rsget("itemoption")
						flist(i).fitemname		= db3_rsget("itemname")
						flist(i).foptionname	= db3_rsget("optionname")
						flist(i).fyyyymmdd		= db3_rsget("yyyymmdd")
						flist(i).fmakerID		= db3_rsget("makerID")
						flist(i).fcdL			= db3_rsget("cdL")
						flist(i).fcdM			= db3_rsget("cdM")
						flist(i).fcdS			= db3_rsget("cdS")
						flist(i).fsaleNo		= db3_rsget("saleNo")
						flist(i).fsaleCost		= db3_rsget("saleCost")
						flist(i).forderCnt		= db3_rsget("orderCnt")
						flist(i).forderManCnt	= db3_rsget("orderManCnt")
						flist(i).forderWomanCnt	= db3_rsget("orderWomanCnt")
						flist(i).forderManAge	= db3_rsget("orderManAge")
						flist(i).forderWomanAge	= db3_rsget("orderWomanAge")
						flist(i).fpassday		= db3_rsget("passday")
						flist(i).fcatename		= db3_rsget("cdLName")
						
					i=i+1
					db3_rsget.moveNext
				Loop
			End If
			
			db3_rsget.close
		END IF
	end function
	
	
	
	public function fbaesong_graph		'배송 소요일 분석 그래프
	dim i , sql, strSubSql
		
		If FSDate <> "" AND FEDate <> "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd Between '" & FSDate & "' AND '" & FEDate & "' "
		ElseIf FSDate <> "" AND FEDate = "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd >= '" & FSDate & "' "
		ElseIf FSDate = "" AND FEDate <> "" Then
			strSubSql = strSubSql & " AND m.yyyymmdd <= '" & FEDate & "' "
		End If
		
		If FIsNotZero = "Y" Then
			strSubSql = strSubSql & " AND m.passday > '0' "
		End If

		If FItemID <> "" Then
			strSubSql = strSubSql & " AND m.itemID = '" & FItemID & "' "
		End If
	
		sql = "SELECT " & _
			  "			Year(yyyymmdd) AS yyyy, Month(yyyymmdd) AS mm, " & _
			  "			m.itemid, (sum(Convert(float,m.passday))/sum(Convert(float,m.pdCNT))) AS delayday, " & _
			  "			(SELECT itemname From [" & FDBName & "].[dbo].[tbl_item] WHERE itemid = m.itemid) AS itemname " & _
			  "		FROM " & _
			  "			[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
			  "		WHERE " & _
			  "			1=1 " & strSubSql & " " & _
			  "		GROUP BY year(yyyymmdd), month(yyyymmdd), m.itemid " & _
			  "		ORDER BY year(yyyymmdd) ASC, month(yyyymmdd) ASC " & _
			  ""
		
		'response.write sql&"<br>"	
		db3_rsget.open sql,db3_dbget,1	
		
		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set flist(i) = new Cbaesong_oneitem
					
					flist(i).fyyyy	 		= db3_rsget("yyyy")
					flist(i).fmm	 		= db3_rsget("mm")
					flist(i).fitemid 		= db3_rsget("itemid")
					flist(i).fdelayday 		= db3_rsget("delayday")
					flist(i).fitemname		= db3_rsget("itemname")
					
			db3_rsget.movenext
			i = i + 1
			loop	
		end if
		db3_rsget.close

	end function	


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


Function CategorySelectBox(vGubun,vValue)
	Dim sql, vDBName, vSelect
	
	IF application("Svr_Info") = "Dev" THEN
		vDBName	= "db_back"
	Else
		vDBName	= "db_datamart"
	End If
		
	If vGubun = "large" Then
		Response.Write "<select name='cate_large' class='select'>"
		Response.Write "<option value=''>전체</option>"
		
		sql = "SELECT code_large, code_nm FROM [" & vDBName & "].[dbo].[tbl_Cate_large] WHERE display_yn = 'Y'"
		db3_rsget.open sql,db3_dbget,1
		Do Until db3_rsget.Eof
			If CStr(vValue) = CStr(db3_rsget("code_large")) Then
				vSelect = "selected"
			End If
			Response.Write "<option value='" & db3_rsget("code_large") & "' " & vSelect & ">" & db3_rsget("code_nm") & "</option>"
			vSelect = ""
		db3_rsget.MoveNext
		Loop
		db3_rsget.close
		
		Response.Write "</select>"
	End If
End Function











'	public function fbaesong_groupbylist		'배송지연 리스트
'	dim i , sql, strSubSql
'		
'		If FSDate <> "" AND FEDate <> "" Then
'			strSubSql = strSubSql & " AND m.yyyymmdd Between '" & FSDate & "' AND '" & FEDate & "' "
'		ElseIf FSDate <> "" AND FEDate = "" Then
'			strSubSql = strSubSql & " AND m.yyyymmdd >= '" & FSDate & "' "
'		ElseIf FSDate = "" AND FEDate <> "" Then
'			strSubSql = strSubSql & " AND m.yyyymmdd <= '" & FEDate & "' "
'		End If
'		
'		If FIsNotZero = "Y" Then
'			strSubSql = strSubSql & " AND m.passday > '0' "
'		End If
'		
'		If FGubun = "i" Then	'####### 상품별 group by
'			'If FItemID <> "" Then
'			'	strSubSql = strSubSql & " AND m.itemID = '" & FItemID & "' "
'			'End If
'			
'			sql = "SELECT COUNT(*) FROM (" & _
'				  "		SELECT COUNT(*) as cnt FROM [db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
'				  "		WHERE 1=1 " & strSubSql & " GROUP BY m.itemid " & _
'				  ") AS A "
'			db3_rsget.open sql,db3_dbget,1
'			'response.write sql
'			IF not db3_rsget.EOF THEN
'				FTotalCount = db3_rsget(0)
'			END IF
'			db3_rsget.close
'					
'			IF FTotalCount > 0 THEN
'				sql = "SELECT TOP "& (FPageSize * FCurrPage) &" " & _
'					  "			m.itemid, (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) AS delayday, " & _
'					  "			(SELECT itemname From [" & FDBName & "].[dbo].[tbl_item] WHERE itemid = m.itemid) AS itemname " & _
'					  "		FROM " & _
'					  "			[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
'					  "		WHERE " & _
'					  "			1=1 " & strSubSql & " " & _
'					  "		GROUP BY m.itemid " & _
'					  "		ORDER BY (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) DESC " & _
'					  ""
'				
'				'response.write sql&"<br>"	
'				db3_rsget.open sql,db3_dbget,1	
'				
'				if (FCurrPage * FPageSize < FTotalCount) then
'					FResultCount = FPageSize
'				else
'					FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
'				end if
'				
'				FTotalPage = (FTotalCount\FPageSize)
'				
'				if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
'				
'				redim preserve flist(FResultCount)
'	
'				db3_rsget.PageSize= FPageSize
'				If  not db3_rsget.EOF  then
'					db3_rsget.absolutepage = FCurrPage
'					Do Until db3_rsget.Eof
'						set flist(i) = new Cbaesong_oneitem
'
'							flist(i).fitemid 		= db3_rsget("itemid")
'							flist(i).fdelayday 		= db3_rsget("delayday")
'							flist(i).fitemname		= db3_rsget("itemname")
'							
'					db3_rsget.movenext
'					i = i + 1
'					loop	
'				end if
'				db3_rsget.close
'			End If
'		ElseIf FGubun = "io" Then	'####### 상품,옵션별 group by
'			'If FItemID <> "" Then
'			'	strSubSql = strSubSql & " AND m.itemID = '" & FItemID & "' "
'			'End If
'		
'				sql = "SELECT Top " & FTop & " " & _
'					  "			m.itemid, m.itemoption, (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) AS delayday, " & _
'					  "			(SELECT itemname From [" & FDBName & "].[dbo].[tbl_item] WHERE itemid = m.itemid) AS itemname, " & _
'					  "			(SELECT optionname From [" & FDBName & "].[dbo].[tbl_item_option] WHERE itemid = m.itemid AND itemoption = m.itemoption) AS optionname " & _
'					  "		FROM " & _
'					  "			[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
'					  "		WHERE " & _
'					  "			1=1 " & strSubSql & " " & _
'					  "		GROUP BY m.itemid, m.itemoption " & _
'					  "		ORDER BY (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) DESC " & _
'					  ""
'				
'				'response.write sql&"<br>"	
'				db3_rsget.open sql,db3_dbget,1	
'				
'				FTotalCount = db3_rsget.recordcount
'				redim flist(FTotalCount)
'				i = 0
'				if not db3_rsget.eof then
'					do until db3_rsget.eof
'						set flist(i) = new Cbaesong_oneitem
'							
'							flist(i).fitemid 		= db3_rsget("itemid")
'							flist(i).fitemoption 	= db3_rsget("itemoption")
'							flist(i).fdelayday 		= db3_rsget("delayday")
'							flist(i).fitemname		= db3_rsget("itemname")
'							flist(i).foptionname	= db3_rsget("optionname")
'							
'					db3_rsget.movenext
'					i = i + 1
'					loop	
'				end if
'				db3_rsget.close
'		ElseIf FGubun = "m" Then	'####### 브랜드별 group by
'			'If FMakerID <> "" Then
'			'	strSubSql = strSubSql & " AND m.makerID = '" & FMakerID & "' "
'			'End If
'			
'				sql = "SELECT Top " & FTop & " " & _
'					  "			m.makerID, (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) AS delayday " & _
'					  "		FROM " & _
'					  "			[db_datamart].[dbo].[tbl_mkt_daily_itemsale_sellDate] AS m " & _
'					  "		WHERE " & _
'					  "			1=1 " & strSubSql & " " & _
'					  "		GROUP BY m.makerID " & _
'					  "		ORDER BY (sum(Convert(float,m.passday))/sum(Convert(float,m.orderCNT))) DESC " & _
'					  ""
'				
'				'response.write sql&"<br>"	
'				db3_rsget.open sql,db3_dbget,1	
'				
'				FTotalCount = db3_rsget.recordcount
'				redim flist(FTotalCount)
'				i = 0
'				if not db3_rsget.eof then
'					do until db3_rsget.eof
'						set flist(i) = new Cbaesong_oneitem
'							
'							flist(i).fmakerID 		= db3_rsget("makerID")
'							flist(i).fdelayday 		= db3_rsget("delayday")
'							
'					db3_rsget.movenext
'					i = i + 1
'					loop	
'				end if
'				db3_rsget.close
'		End If
'	end function	
%>