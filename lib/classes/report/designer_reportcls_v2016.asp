<%
Class CSellReport

	public FSPageNo
	public FEPageNo
	public FPageSize
	public FCurrPage
	public FTotCnt
	
	public FRectDispCate
	public FRectDateGijun
	public FRectSort
	public FRectMakerid
	public FRectStartdate
	public FRectEndDate
	public FRectTerm
	public FRectIsOption
	public FRectItemid
	
	public FRectIsOpt
	public FRectckpointsearch
	public FRectRegStart
	public FRectRegEnd 

	public Function fnGetSellReport
		Dim strSql, strOrder, strOrder1
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		Dim searchStDt : searchStDt = LEFT(dateAdd("d",-365,NOW()),10)

		strSql =" select  count(*) FROM ("
		strSql = strSql & " select "
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  as ddate " & vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  as ddate " & vbCRLF
		end if   
 
		strSql = strSql & "  FROM  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN [db_statistics].dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'" & vbCRLF
		END IF
		
		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'" & vbCRLF
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) " & vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' " & vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF

		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")" & vbCRLF
		END IF
		
		strSql = strSql & "	GROUP BY "
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  "& vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  "& vbCRLF
		end if 
		
 		strSql = strSql & " ) AS TB "

 		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			FTotCnt = rsSTSget(0)
		End If 
		rsSTSget.close
 
	if FTotCnt > 0 THEN
		dim ddate 
		if  FRectTerm ="m" then
			ddate =  " convert(varchar(7),d."&FRectDateGijun&",121)  "
		else
			ddate = " convert(varchar(10),d."&FRectDateGijun&",121)  "
		end if 
 	 
  
		if FRectSort = "DA" THEN
			strOrder= ddate & " asc  "  & vbCRLF
		elseif  FRectSort = "CD" THEN
			strOrder= " sum(d.itemno) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "CA" THEN
			strOrder= " sum(d.itemno) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MD" THEN
			strOrder= " sum(d.itemcost) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MA" THEN
			strOrder= " sum(d.itemcost) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BD" THEN
			strOrder= " sum(d.buycash) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BA" THEN
			strOrder= " sum(d.buycash) asc, "&ddate&" desc "& vbCRLF
		else 
			strOrder= ddate & " desc  "	& vbCRLF		  
		end if
  
		if FRectSort = "DA" THEN
			strOrder1=  " ddate asc  "  & vbCRLF
		elseif  FRectSort = "CD" THEN
			strOrder1= " itemno desc, ddate desc "& vbCRLF
		elseif  FRectSort = "CA" THEN
			strOrder1= " itemno asc, ddate desc "& vbCRLF
		elseif  FRectSort = "MD" THEN
			strOrder1= " itemcost desc, ddate desc "& vbCRLF
		elseif  FRectSort = "MA" THEN
			strOrder1= " itemcost asc, ddate desc "& vbCRLF
		elseif  FRectSort = "BD" THEN
			strOrder1= " buycash desc, ddate desc "& vbCRLF
		elseif  FRectSort = "BA" THEN
			strOrder1= " buycash asc, ddate desc "& vbCRLF
		else 
			strOrder1= " ddate desc  "& vbCRLF
		end if
	
		strSql = " SELECT TB.* from ( select " 
		strSql = strSql & " 	ROW_NUMBER() OVER (ORDER BY " & strOrder & " ) as RowNum, "& vbCRLF
		strSql = strSql & "		sum(d.itemno) AS itemno, "
		strSql = strSql & "		sum(d.orgitemcost*d.itemno) AS orgitemcost, "
		strSql = strSql & "		sum(d.itemcostCouponNotApplied*d.itemno) AS itemcostCouponNotApplied, "
		strSql = strSql & "		sum(d.itemcost*d.itemno) AS itemcost, "
		strSql = strSql & "		sum(d.buycash*d.itemno) as buycash, "
		strSql = strSql & "		sum(d.reducedPrice*d.itemno) as reducedprice "		
		if  FRectTerm ="m" then
			strSql = strSql & "	,convert(varchar(7),d."&FRectDateGijun&",121)  as ddate "& vbCRLF
		else
			strSql = strSql & "	,convert(varchar(10),d."&FRectDateGijun&",121)  as ddate  "& vbCRLF
		end if   
 
		strSql = strSql & "  From  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"& vbCRLF
		END IF
		
 		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"& vbCRLF
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) "& vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' "& vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF
		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")"& vbCRLF
		END IF
		strSql = strSql & "	GROUP BY "
	
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  "& vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  "& vbCRLF
		end if
		strSql = strSql & " ) as TB "
		strSql = strSql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo& vbCRLF
		strSql = strSql & " order by " & strOrder1& vbCRLF

		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			fnGetSellReport = rsSTSget.getRows()
		End If 
		rsSTSget.close
	END IF
	End Function

	public Function fnGetSellReportCSV
		Dim strSql, strOrder, strOrder1
	  	
		Dim searchStDt : searchStDt = LEFT(dateAdd("d",-365,NOW()),10)

 		dim ddate 
 	 	 if  FRectTerm ="m" then
			ddate = " convert(varchar(7),d."&FRectDateGijun&",121) "& vbCRLF
		else
			ddate = " convert(varchar(10),d."&FRectDateGijun&",121) "& vbCRLF
		end if 
 	 
		
		if FRectSort = "DA" THEN
			strOrder= ddate & " asc  "  & vbCRLF
		elseif  FRectSort = "CD" THEN
			strOrder= " sum(d.itemno) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "CA" THEN
			strOrder= " sum(d.itemno) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MD" THEN
			strOrder= " sum(d.itemcost) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MA" THEN
			strOrder= " sum(d.itemcost) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BD" THEN
			strOrder= " sum(d.buycash) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BA" THEN
			strOrder= " sum(d.buycash) asc, "&ddate&" desc "& vbCRLF
		else 
			strOrder= ddate & " desc  "& vbCRLF
		end if
   
	
		strSql = " select top " &FPageSize &" ROW_NUMBER() OVER (ORDER BY " & strOrder & " ) as RowNum,"& vbCRLF
		strSql = strSql & "		sum(d.itemno) AS itemno, "
		strSql = strSql & "		sum(d.orgitemcost*d.itemno) AS orgitemcost, "
		strSql = strSql & "		sum(d.itemcostCouponNotApplied*d.itemno) AS itemcostCouponNotApplied, "
		strSql = strSql & "		sum(d.itemcost*d.itemno) AS itemcost, "
		strSql = strSql & "		sum(d.buycash*d.itemno) as buycash, "
		strSql = strSql & "		sum(d.reducedPrice*d.itemno) as reducedprice "		
		if  FRectTerm ="m" then
			strSql = strSql & "	,convert(varchar(7),d."&FRectDateGijun&",121)  as ddate "& vbCRLF
		else
			strSql = strSql & "	,convert(varchar(10),d."&FRectDateGijun&",121)  as ddate  "& vbCRLF
		end if   
  
		strSql = strSql & "  From  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"& vbCRLF
		END IF
		
 		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"& vbCRLF
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) "& vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' "& vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF
		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")"& vbCRLF
		END IF
		strSql = strSql & "	GROUP BY "
	
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  "& vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  "& vbCRLF
		end if  
 
		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			fnGetSellReportCSV = rsSTSget.getRows()
		End If 
		rsSTSget.close
 
	End Function
	
'--==========================================================================================================

	public Function fnGetSellItemReport
		Dim strSql, strOrder, strOrder1
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		Dim searchStDt : searchStDt = LEFT(dateAdd("d",-365,NOW()),10)

		strSql =" select  count(*) FROM ("
		strSql = strSql & " select "
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  as ddate " & vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  as ddate  " & vbCRLF
 		end if   
 		 
		strSql = strSql & "	,d.itemid"
		if FRectIsOption = "Y" then
			strSql = strSql & ", d.itemoption "
		end if		
		strSql = strSql & "  FROM  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
		
		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'" & vbCRLF
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) " & vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' " & vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF

		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")" & vbCRLF
		END IF

		strSql = strSql & "	GROUP BY "
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  " & vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  " & vbCRLF
 		end if   
		 strSql = strSql & "	, d.itemid"
 		if FRectIsOption = "Y" then
			strSql = strSql & ", d.itemoption "
		end if		
 		strSql = strSql & " ) AS TB "
 
 		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			FTotCnt = rsSTSget(0)
		End If 
		rsSTSget.close
 
 	if FTotCnt > 0 THEN
 		dim ddate 
 	 	if  FRectTerm ="m" then
			ddate =  "	convert(varchar(7),d."&FRectDateGijun&",121)  " & vbCRLF
		else
			ddate = " convert(varchar(10),d."&FRectDateGijun&",121)  " & vbCRLF
 		end if 
 	 
  
		if FRectSort = "DA" THEN
			strOrder= ddate & " asc  , d.itemid desc "  & vbCRLF
		elseif  FRectSort = "TD" THEN
			strOrder= "d.itemid desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "TA" THEN
			strOrder= "d.itemid asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "CD" THEN
			strOrder= " sum(d.itemno) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "CA" THEN
			strOrder= " sum(d.itemno) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MD" THEN
			strOrder= " sum(d.itemcost*d.itemno) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "MA" THEN
			strOrder= " sum(d.itemcost*d.itemno) asc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BD" THEN
			strOrder= " sum(d.buycash*d.itemno) desc, "&ddate&" desc "& vbCRLF
		elseif  FRectSort = "BA" THEN
			strOrder= " sum(d.buycash*d.itemno) asc, "&ddate&" desc "& vbCRLF
		else 
			strOrder= ddate & " desc  , d.itemid desc " & vbCRLF
		end if
  		if FRectIsOption = "Y" then
			 strOrder = strOrder & ", d.itemoption "& vbCRLF
		end if
  
		if FRectSort = "DA" THEN
			strOrder1=  " ddate asc  , itemid desc " & vbCRLF
		elseif  FRectSort = "TD" THEN
			strOrder1= " itemid desc, ddate desc "& vbCRLF
		elseif  FRectSort = "TA" THEN
			strOrder1= " itemid asc, ddate desc "& vbCRLF
		elseif  FRectSort = "CD" THEN
			strOrder1= " itemno desc, ddate desc "& vbCRLF
		elseif  FRectSort = "CA" THEN
			strOrder1= " itemno asc, ddate desc "& vbCRLF
		elseif  FRectSort = "MD" THEN
			strOrder1= " itemcost desc, ddate desc "& vbCRLF
		elseif  FRectSort = "MA" THEN
			strOrder1= " itemcost asc, ddate desc "& vbCRLF
		elseif  FRectSort = "BD" THEN
			strOrder1= " buycash desc, ddate desc "& vbCRLF
		elseif  FRectSort = "BA" THEN
			strOrder1= " buycash asc, ddate desc "& vbCRLF
		else 
			strOrder1= " ddate desc   , itemid desc " & vbCRLF
		end if

		if FRectIsOption = "Y" then
			 strOrder1 = strOrder1 & ", itemoption "& vbCRLF
		end if
		
		strSql = " SELECT TB.RowNum, TB.itemno, TB.orgitemcost, TB.itemcostCouponNotApplied"& vbCRLF
		strSql = strSql & ",TB.itemcost,TB.buycash,TB.reducedprice,TB.ddate,TB.itemid"& vbCRLF
		strSql = strSql & ",i.smallimage,  i.itemname"& vbCRLF
		if FRectIsOption = "Y" then
			strSql = strSql & ", TB.itemoption,isNULL(op.optionname,'') as optionname "& vbCRLF
		end if
		strSql = strSql & "	from ( select " 
		strSql = strSql & " 	ROW_NUMBER() OVER (ORDER BY " & strOrder & " ) as RowNum, "& vbCRLF
		strSql = strSql & "		sum(d.itemno) AS itemno, "
		strSql = strSql & "		sum(d.orgitemcost*d.itemno) AS orgitemcost, "
		strSql = strSql & "		sum(d.itemcostCouponNotApplied*d.itemno) AS itemcostCouponNotApplied, "
		strSql = strSql & "		sum(d.itemcost*d.itemno) AS itemcost, "
		strSql = strSql & "		sum(d.buycash*d.itemno) as buycash, "
		strSql = strSql & "		sum(d.reducedPrice*d.itemno) as reducedprice "		
		if  FRectTerm ="m" then
			strSql = strSql & "	,convert(varchar(7),d."&FRectDateGijun&",121)  as ddate "& vbCRLF
		else
			strSql = strSql & "	,convert(varchar(10),d."&FRectDateGijun&",121)  as ddate  "& vbCRLF
		end if   
 
		strSql = strSql & "	,d.itemid"
		if FRectIsOption = "Y" then
			strSql = strSql & ", d.itemoption "
		end if
		strSql = strSql & "  From  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"& vbCRLF
		END IF
		
 		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"& vbCRLF
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) "& vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' "& vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF

		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")"& vbCRLF
		END IF
		strSql = strSql & "	GROUP BY "
	
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  "& vbCRLF
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  "& vbCRLF
		end if   
		strSql = strSql & "	, d.itemid"
 		if FRectIsOption = "Y" then
			strSql = strSql & ", d.itemoption "
		end if
		strSql = strSql & " ) as TB "
		strSql = strSql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i ON TB.itemid = i.itemid "
		if FRectIsOption = "Y" then
			strSql = strSql & " LEFT OUTER JOIN db_statistics.dbo.tbl_item_option as op on TB.itemid = op.itemid  and TB.itemoption = op.itemoption "
		end if
		strSql = strSql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo& vbCRLF
		strSql = strSql & " order by " & strOrder1  

		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			fnGetSellItemReport = rsSTSget.getRows()
		End If 
			rsSTSget.close
		END IF
	End Function



	public Function fnGetSellItemReportCSV
		Dim strSql, strOrder, strOrder1
	  
	  	Dim searchStDt : searchStDt = LEFT(dateAdd("d",-365,NOW()),10)
		  
 		dim ddate 
		if  FRectTerm ="m" then
			ddate =  "	convert(varchar(7),d."&FRectDateGijun&",121)  "
		else
			ddate= "	convert(varchar(10),d."&FRectDateGijun&",121)  "
		end if 
 	 
  
		if FRectSort = "DA" THEN
			strOrder= ddate & " asc   , d.itemid desc " 
		elseif  FRectSort = "TD" THEN
			strOrder= "d.itemid desc, "&ddate&" desc "
		elseif  FRectSort = "TA" THEN
			strOrder= "d.itemid asc, "&ddate&" desc "
		elseif  FRectSort = "CD" THEN
			strOrder= " sum(d.itemno) desc, "&ddate&" desc "
		elseif  FRectSort = "CA" THEN
			strOrder= " sum(d.itemno) asc, "&ddate&" desc "
		elseif  FRectSort = "MD" THEN
			strOrder= " sum(d.itemcost) desc, "&ddate&" desc "
		elseif  FRectSort = "MA" THEN
			strOrder= " sum(d.itemcost) asc, "&ddate&" desc "
		elseif  FRectSort = "BD" THEN
			strOrder= " sum(d.buycash) desc, "&ddate&" desc "
		elseif  FRectSort = "BA" THEN
			strOrder= " sum(d.buycash) asc, "&ddate&" desc "
		else 
			strOrder= ddate & " desc   , d.itemid desc " 
		end if
		if FRectIsOption = "Y" then
			strOrder= strOrder & ", op.itemoption "
		end if
		strSql = " select top " &FPageSize &" ROW_NUMBER() OVER (ORDER BY " & strOrder & " ) as RowNum,"
		strSql = strSql & "		sum(d.itemno) AS itemno, "
		strSql = strSql & "		sum(d.orgitemcost*d.itemno) AS orgitemcost, "
		strSql = strSql & "		sum(d.itemcostCouponNotApplied*d.itemno) AS itemcostCouponNotApplied, "
		strSql = strSql & "		sum(d.itemcost*d.itemno) AS itemcost, "
		strSql = strSql & "		sum(d.buycash*d.itemno) as buycash, "
		strSql = strSql & "		sum(d.reducedPrice*d.itemno) as reducedprice "		
		if  FRectTerm ="m" then
			strSql = strSql & "	,convert(varchar(7),d."&FRectDateGijun&",121)  as ddate "
		else
			strSql = strSql & "	,convert(varchar(10),d."&FRectDateGijun&",121)  as ddate  "
		end if   
		strSql = strSql & "	 ,d.itemid,i.smallimage,  i.itemname" 
		if FRectIsOption = "Y" then
				strSql = strSql & ", op.itemoption, op.optionname "
		end if
		strSql = strSql & "  From  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		strSql = strSql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
		IF FRectDispCate<>"" THEN	 
			strSql = strSql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
		if FRectIsOption = "Y" then
			  strSql = strSql & " LEFT OUTER JOIN db_statistics.dbo.tbl_item_option as op on d.itemid = op.itemid  and d.itemoption = op.itemoption "
		end if
		
 		strSql = strSql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		strSql = strSql & " AND d.ipkumdate is Not NULL"
		strSql = strSql & " AND d.cancelyn='N'"
		strSql = strSql & " AND d.dcancelyn<>'Y'"
		strSql = strSql & " AND d.itemid not in (0,100) "& vbCRLF
		strSql = strSql & " AND d.beadaldiv not in (90)" & vbCRLF
		strSql = strSql & " AND d.makerid = '"&FRectMakerid&"' "& vbCRLF
		strSql = strSql & " AND d."&FRectDateGijun&">='"&searchStDt&"'" & vbCRLF
		IF FRectItemid <> "" Then
			strSql = strSql & " and d.itemid in ("& FRectItemID&")"
		END IF
		strSql = strSql & "	GROUP BY "
	
		if  FRectTerm ="m" then
			strSql = strSql & "	convert(varchar(7),d."&FRectDateGijun&",121)  "
		else
			strSql = strSql & "	convert(varchar(10),d."&FRectDateGijun&",121)  "
		end if   
 		strSql = strSql & "	, d.itemid,i.smallimage,  i.itemname" 
 		if FRectIsOption = "Y" then
			strSql = strSql & ", op.itemoption, op.optionname "
		end if
	 
		'strSql = strSql & " order by " & strOrder
	  
		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			fnGetSellItemReportCSV = rsSTSget.getRows()
		End If 
		rsSTSget.close
 
	End Function
	
	'--==========================================================================================================

	public Function fnGetBestSeller
	 	dim sqlStr

		Dim searchStDt : searchStDt = LEFT(dateAdd("d",-365,NOW()),10)

		sqlStr = "select top " & CStr(FPageSize)
		sqlStr = sqlStr & " d.itemid , "
		sqlStr = sqlStr & " i.itemname, d.makerid "
		sqlStr = sqlStr & " , sum(d.itemno) as sm ,sum(d.buycash*d.itemno)as sm2, sum(d.itemcost*d.itemno) as sm3 "
		sqlStr = sqlStr & ", i.smallimage "
		if FRectIsOpt ="Y"	Then	 
			sqlStr = sqlStr & ", o.optionname"
		end if		
		sqlStr = sqlStr & "  FROM  [db_statistics_order].dbo.tbl_order_detail_raw  d " & vbCRLF
		sqlStr = sqlStr & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
		if FRectIsOpt ="Y" Then
		sqlStr = sqlStr & "		left outer join [db_statistics].dbo.tbl_item_option as o on  d.itemid = o.itemid and d.itemoption = o.itemoption  "
		end if
		IF FRectDispCate<>"" THEN	 
			sqlStr = sqlStr & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"&vbCrlf
		END IF
		sqlStr = sqlStr & " where d.makerid= '" & FRectMakerid & "'"&vbCrlf
		sqlStr = sqlStr & " AND d.ipkumdate is Not NULL"
		sqlStr = sqlStr & " AND d.cancelyn='N'"
		sqlStr = sqlStr & " AND d.dcancelyn<>'Y'"
		sqlStr = sqlStr & " AND d.itemid not in (0,100) " & vbCRLF
		sqlStr = sqlStr & " AND d.beadaldiv not in (90)" & vbCRLF
		sqlStr = sqlStr & " AND d.regdate>='"&searchStDt&"'" & vbCRLF
		if (FRectStartdate<>"") then
			sqlStr = sqlStr & " and d.regdate >='" & CStr(FRectStartdate) & "'"&vbCrlf
		end if

		if (FRectEndDate<>"") then
			sqlStr = sqlStr & " and d.regdate <'" & CStr(DateAdd("d",1,GetValidDate(FRectEndDate))) & "'"&vbCrlf
		end if 

		''없음.
		''if (FRectckpointsearch = "") then
		''	sqlStr = sqlStr & " and m.accountdiv <> 30"
		''end if
		sqlStr = sqlStr & " group by d.itemid,    i.itemname, d.makerid, i.smallimage "
		if FRectIsOpt ="Y"	Then	 
			sqlStr = sqlStr & ", o.optionname"
		end if	
		if FRectSort = "CD"	 Then '상품수
			sqlStr = sqlStr & " order by sm Desc"  		
		elseif 	FRectSort = "CA" Then  
			sqlStr = sqlStr & " order by sm asc"  		
		elseif 	FRectSort = "BD" Then '정산액
			sqlStr = sqlStr & " order by sm2 Desc"  
		elseif 	FRectSort = "BA" Then '정산액
			sqlStr = sqlStr & " order by sm2 asc"  		
		elseif 	FRectSort = "MA" Then '정산액
			sqlStr = sqlStr & " order by sm3 asc" 		
		else		
			sqlStr = sqlStr & " order by sm3 Desc"  		
		end if		
  
		rsSTSget.CursorLocation = adUseClient
    	rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly
    
		If Not rsSTSget.Eof Then
			FTotCnt =  rsSTSget.recordCount
			fnGetBestSeller = rsSTSget.getRows()
		End If 
		rsSTSget.close 
	 
	End Function 
End Class


%>