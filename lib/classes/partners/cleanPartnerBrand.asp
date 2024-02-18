<%
class CCleanItem

public FTotCnt
public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage

public fDBDATAMART
public fDBSELFORDER
public fDBSELFITEM

Private Sub Class_Initialize()
IF application("Svr_Info")="Dev" THEN
	fDBDATAMART="TENDB"
else
	fDBDATAMART="DBDATAMART"
end if
IF application("Svr_Info")="Dev" THEN
	fDBSELFORDER="TENDB.db_order"
else
	fDBSELFORDER="db_analyze_data_raw"
end if

IF application("Svr_Info")="Dev" THEN
	fDBSELFITEM="TENDB.db_item"
else
	fDBSELFITEM="db_analyze_data_raw"
end if
End Sub		

 
public FRectStdate
public FRectEddate
public FRectWishCount
public FRectIsUsing
public FRectMakerid
public FRectDispCate
public FRectSellYN
public FRectMwdiv
public FRectSort
public FRectsocname_kr
public FRectcrect     
public FRectcompanyno 
public FRectgroupid   

    public Function fnGetCleanBrandList
    if FRectWishCount ="" then FRectWishCount = 5
    if FRectStdate ="" then FRectStdate = dateadd("m",-6,date()) 
    if FRectEddate ="" then FRectEddate = dateadd("m",3,FRectStdate) 
     
    dim strSql
      
    dim strSqlAdd, strSqlSort, strSqlAdd1,strSqlSort1
    strSqlAdd = "" 
    strSqlAdd1 = ""

	If FRectMakerid <> "" Then
	    strSqlAdd = strSqlAdd & " and i.makerid = '" & FRectMakerid &"'"
	end if
 
    
    IF FRectSellYN <> "" Then
        if FRectSellYN = "YS" THEN
		    strSqlAdd = strSqlAdd & " and i.sellyn <> 'N'"
	    else
	        strSqlAdd = strSqlAdd & " and i.sellyn = '"&FRectSellYN&"'"
        end if
	END IF
	 

    strSql = " select i.makerid, sum(c.favcount) as sumfavcount , count(i.itemid) as itemcount "
    strSql = strSql & " into #tmpWish "
    strSql = strSql & "FROM  db_analyze_data_raw.dbo.tbl_item as i "
    strSql = strSql & "inner join  db_analyze_data_raw.dbo.tbl_item_Contents   as c on i.itemid = c.itemid  "
    IF FRectDispCate<>"" THEN
			strSql = strSql & " 	inner JOIN "& fDBSELFITEM &".dbo.tbl_display_cate_item as dc"
			strSql = strSql & " 		on i.itemid = dc.itemid  and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
	END IF
    strSql = strSql & " where  i.isusing = 'Y' "
        if FRectMwDiv ="MW" then
            strSql = strSql & " and (i.mwdiv = 'M' or i.mwdiv='W') " 
        elseif FRectMwDiv <> "" then    
	        strSql = strSql & " and i.mwdiv = '" & FRectMwDiv &"'"
	    else
		    strSql = strSql & " and i.mwdiv = 'U'"
	    end if  
     strSql = strSql &  strSqlAdd
     strSql = strSql & " group by i.makerid having sum(c.favcount) < "&FRectWishCount &Vbcrlf
    
     
    strSql = strSql & " select w.makerid, w.sumfavcount, w.itemcount  "
    strSql = strSql & " into #tmpOrder "
    strSql = strSql & " from #tmpWish as w "
    strSql = strSql & "left outer join db_analyze_data_raw.dbo.tbl_order_detail as d on w.makerid = d.makerid and d.cancelyn<>'Y' and d.beasongdate >='"&FRectStdate&"' and d.beasongdate < '"&FRectEddate&"' "
    strSql = strSql & "left outer join db_analyze_data_raw.dbo.tbl_order_master as m on m.orderserial = d.orderserial and m.ipkumdiv>3 and m.cancelyn='N' "
    strSql = strSql & " where d.makerid is null "  
    '  response.write strSql
    dbanalget.Execute strSql
    
    if FRectsocname_kr <> "" then
    strSqlAdd1 = strSqlAdd1& " and c.socname_kor like '%"&FRectsocname_kr&"%' "
  	end if	
  	
  	 if FRectcrect <> "" then
    strSqlAdd1 = strSqlAdd1& " and company_name like '%"&FRectcrect&"%' "
  	end if	
  	
  	 if FRectcompanyno <> "" then
    strSqlAdd1 = strSqlAdd1& " and company_no = '"&FRectcompanyno&"' "
  	end if	
  	
  	 if FRectgroupid <> "" then
    strSqlAdd1 = strSqlAdd1& " and groupid = '"&FRectgroupid&"' "
  	end if	
    
    strSql =   " select count(t.makerid)  "
    strSql =  strSql & " from #tmpOrder as t "
    strSql =  strSql & " inner join db_analyze_data_raw.[dbo].[tbl_partner] as p on p.id = t.makerid " 
    strSql =  strSql & " inner join dbdatamart.db_user.dbo.tbl_user_c as c on p.id = c.userid "
    strSql =  strSql & " where p.isusing = 'Y' and c.isusing ='Y' " & strSqlAdd1      
    
    rsAnalget.open strSql,dbAnalget,1
    	IF Not (rsAnalget.EOF OR rsAnalget.BOF) THEN
    	    FTotCnt = rsAnalget(0)
        END IF
     rsAnalget.close 
    IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

'		IF FRectSort = "ID" THEN
'			strSqlSort= " t.itemid desc, i.regdate asc, i.sellstdate asc "
'		ELSEIF	FRectSort = "RD" THEN
'			strSqlSort= " i.regdate desc ,t.itemid Asc, i.sellstdate asc "
'		ELSEIF	FRectSort = "RA" THEN
'			strSqlSort= " i.regdate asc, t.itemid Asc, i.sellstdate asc "
'		ELSEIF	FRectSort = "SD" THEN
'			strSqlSort= "  i.sellstdate desc ,t.itemid Asc, i.regdate asc"
'		ELSEIF	FRectSort = "SA" THEN
'			strSqlSort= "  i.sellstdate asc,t.itemid Asc, i.regdate asc "
'		ELSE
'			strSqlSort= " t.itemid Asc, i.regdate asc, i.sellstdate asc "
'		END IF				

		strSqlSort =" p.regdate "
		strSqlSort1=" regdate "
		
    strSql =  " SELECT   makerid, socname_kor, socname,  groupid, company_no, company_name,regdate, sumfavcount, itemcount, isoffusing,isextusing  "	
 	strSql =  strSql & " from ( "
    strSql =  strSql & " select ROW_NUMBER() OVER (ORDER BY  "&strSqlSort&" ) as RowNum "
    strSql =  strSql & " , t.makerid, c.socname_kor, socname,  groupid, company_no, company_name,p.regdate, t.sumfavcount, t.itemcount  "
    strSql =  strSql & " , isoffusing ,isextusing"
    strSql =  strSql & " from #tmpOrder as t "
    strSql =  strSql & " inner join db_analyze_data_raw.dbo.[tbl_partner] as p on p.id =  t.makerid " 
    strSql =  strSql & " inner join dbdatamart.db_user.dbo.tbl_user_c as c on p.id = c.userid "
    strSql =  strSql & " where p.isusing = 'Y' and c.isusing ='Y'" & strSqlAdd1
    strSql =  strSql & ") AS TB " 
	strSql =  strSql &" WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo   
    strSql =  strSql &" order by  " &strSqlSort1
    response.write strSql
    rsAnalget.open strSql,dbAnalget,1
		IF Not (rsAnalget.EOF OR rsAnalget.BOF) THEN
			fnGetCleanBrandList = rsAnalget.getRows()
		END IF
	rsAnalget.close
	EnD IF	
    End Function
END Class    
%>