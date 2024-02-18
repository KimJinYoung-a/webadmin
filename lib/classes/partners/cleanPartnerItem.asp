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

public FRectItemid
public FRectStdate
public FRectEddate
public FRectWishCount
public FRectIsUsing
public FRectMakerid
public FRectDispCate
public FRectSellYN
public FRectMwdiv
public FRectSort

    public Function fnGetCleanItemList
    if FRectWishCount ="" then FRectWishCount = 5
    if FRectStdate ="" then FRectStdate = dateadd("m",-6,date())
    if FRectEddate ="" then FRectEddate = dateadd("m",3,FRectStdate) 
    dim strSql
'    
'    strSql = " select count(T1.itemid) "
'    strSql =  strSql &"      from ( "
'    strSql =  strSql &"         select  i.itemid  "
'    strSql =  strSql &"        FROM "&fDBSELFITEM&".dbo.tbl_item as i  "
'    strSql =  strSql &"        left outer join  "&fDBDATAMART&".db_my10x10.dbo.tbl_myfavorite as f on i.itemid = f.itemid   "
'    strSql =  strSql &"        group by i.itemid, f.itemid  having count(f.itemid)<5 "
'    strSql =  strSql &"       ) as T1 "
'    strSql =  strSql &"       inner join (  "
'    strSql =  strSql &"      select  i.itemid  "
'    strSql =  strSql &"       from  "&fDBSELFITEM&".dbo.tbl_item as i "
'    strSql =  strSql &"      left outer join  "&fDBSELFORDER&".dbo.tbl_order_detail as d on i.itemid = d.itemid  and d.cancelyn<>'Y' "
'    strSql =  strSql &"      			and d.beasongdate >='2015-11-01' and d.beasongdate < '2016-02-01'  "
'    strSql =  strSql &"     left outer join  "&fDBSELFORDER&".dbo.tbl_order_master as m on m.orderserial = d.orderserial and  m.ipkumdiv>3  and m.cancelyn='N' "
'    strSql =  strSql &"       where d.itemid is null "
'    strSql =  strSql &"       and i.isusing = 'Y'  and i.mwdiv ='U' and i.sellyn <> 'N'"
'    strSql =  strSql &"        ) as T2 on T1.itemid = T2.itemid" 
'    response.write strSql
'response.end
'    rsAnalget.open strSql,dbAnalget,1
'    	IF not rsAnalget.EOF THEN
'    	    FTotCnt = rsAnalget(0)
'        END IF
'       rsAnalget.close 
'    IF FTotCnt > 0 THEN
'		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
'		FEPageNo = FPageSize*FCurrPage
'	strSql = " SELECT itemid "	
'	strSql =  strSql & " from ( "
'    strSql =  strSql & " select ROW_NUMBER() OVER (ORDER BY  T1.itemid ) as RowNum, T1.itemid "
'    strSql =  strSql &"      from ( "
'    strSql =  strSql &"         select  i.itemid  "
'    strSql =  strSql &"        FROM  "&fDBSELFITEM&".dbo.tbl_item as i  "
'    strSql =  strSql &"        left outer join  "&fDBDATAMART&".db_my10x10.dbo.tbl_myfavorite as f on i.itemid = f.itemid   "
'    strSql =  strSql &"        group by i.itemid, f.itemid  having count(f.itemid)<5 "
'    strSql =  strSql &"       ) as T1 "
'    strSql =  strSql &"       inner join (  "
'    strSql =  strSql &"      select  i.itemid  "
'    strSql =  strSql &"       from  "&fDBSELFITEM&".dbo.tbl_item as i "
'    strSql =  strSql &"      left outer join  "&fDBSELFORDER&".dbo.tbl_order_detail as d on i.itemid = d.itemid  and d.cancelyn<>'Y' "
'    strSql =  strSql &"      			and d.beasongdate >='2015-11-01' and d.beasongdate < '2016-02-01'  "
'    strSql =  strSql &"     left outer join  "&fDBSELFORDER&".dbo.tbl_order_master as m on m.orderserial = d.orderserial and  m.ipkumdiv>3  and m.cancelyn='N' "
'    strSql =  strSql &"       where d.itemid is null "
'    strSql =  strSql &"       and i.isusing = 'Y'  and i.mwdiv ='U' and i.sellyn <> 'N'"
'    strSql =  strSql &"        ) as T2 on T1.itemid = T2.itemid"
'    strSql =  strSql &") AS TB " 
'	strSql =  strSql &" WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo   
'    strSql =  strSql &"       order by itemid "
'    
    dim strSqlAdd, strSqlSort,strSqlSort1
    strSqlAdd = "" 
    
	If FRectMakerid <> "" Then
	    strSqlAdd = strSqlAdd & " and i.makerid = '" & FRectMakerid &"'"
	end if
	IF FRectItemid <> "" Then
		strSqlAdd = strSqlAdd & " and i.itemid in ("& FRectItemID&")"
	END IF
    
    IF FRectSellYN <> "" Then
        if FRectSellYN = "YS" THEN
		    strSqlAdd = strSqlAdd & " and i.sellyn <> 'N'"
	    else
	        strSqlAdd = strSqlAdd & " and i.sellyn = '"&FRectSellYN&"'"
        end if
	END IF
	
	 	strSql = " select i.itemid  "
    strSql = strSql & " into #tmpOrder "
    strSql = strSql & "FROM  db_analyze_data_raw.dbo.tbl_item as i "
    strSql = strSql & "left outer join   db_analyze_data_raw.dbo.tbl_order_detail as d on i.itemid = d.itemid and d.cancelyn<>'Y' and d.beasongdate >='"&FRectStdate&"' and d.beasongdate < '"&FRectEddate&"' "
    strSql = strSql & "left outer join   db_analyze_data_raw.dbo.tbl_order_master as m on m.orderserial = d.orderserial  and m.ipkumdiv>3 and m.cancelyn='N' "
    IF FRectDispCate<>"" THEN
			strSql = strSql & " 	inner JOIN "& fDBSELFITEM &".dbo.tbl_display_cate_item as dc"
			strSql = strSql & " 		on i.itemid = dc.itemid  and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
	END IF
    strSql = strSql & " where  d.itemid is null  "
    strSql = strSql & "    and i.isusing = 'Y' and i.regdate < '"&FRectStdate&"'"
        if FRectMwDiv ="MW" then
            strSql = strSql & " and (i.mwdiv = 'M' or i.mwdiv='W') " 
        elseif FRectMwDiv <> "" then    
	        strSql = strSql & " and i.mwdiv = '" & FRectMwDiv &"'"
	    else
		    strSql = strSql & " and i.mwdiv = 'U'"
	    end if  
     strSql = strSql &  strSqlAdd&Vbcrlf
     
    strSql = strSql & " select o.itemid, c.favcount "
    strSql = strSql & " into #tmpWish "
    strSql = strSql & "FROM   #tmpOrder as o "
    strSql = strSql & "inner join  db_analyze_data_raw.dbo.tbl_item_Contents   as c on o.itemid = c.itemid  " 
    strSql = strSql & " where c.favcount< "&FRectWishCount   
    dbanalget.Execute strSql
    
    strSql =   " select count(t.itemid)  "
    strSql =  strSql & " from #tmpWish as t "
    strSql =  strSql & " inner join db_analyze_data_raw.dbo.tbl_item as i on i.itemid = t.itemid "     
    rsAnalget.open strSql,dbAnalget,1
    	IF Not (rsAnalget.EOF OR rsAnalget.BOF) THEN
    	    FTotCnt = rsAnalget(0)
        END IF
     rsAnalget.close 
    IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		IF FRectSort = "ID" THEN
			strSqlSort= " t.itemid desc, i.regdate asc, i.sellstdate asc "
			strSqlSort1="itemid desc,  regdate asc,  sellstdate asc "
		ELSEIF	FRectSort = "RD" THEN
			strSqlSort= " i.regdate desc ,t.itemid Asc, i.sellstdate asc "
			strSqlSort1= "  regdate desc , itemid Asc,  sellstdate asc "
		ELSEIF	FRectSort = "RA" THEN
			strSqlSort= " i.regdate asc, t.itemid Asc, i.sellstdate asc "
			strSqlSort1= "  regdate asc,  itemid Asc,  sellstdate asc "
		ELSEIF	FRectSort = "SD" THEN
			strSqlSort= "  i.sellstdate desc ,t.itemid Asc, i.regdate asc"
			strSqlSort1= "   sellstdate desc , itemid Asc,  regdate asc"
		ELSEIF	FRectSort = "SA" THEN
			strSqlSort= "  i.sellstdate asc,t.itemid Asc, i.regdate asc "
			strSqlSort1= "   sellstdate asc, itemid Asc,  regdate asc "
		ELSE
			strSqlSort= " t.itemid Asc, i.regdate asc, i.sellstdate asc "
			strSqlSort1= "  itemid Asc,  regdate asc,  sellstdate asc "
		END IF				

    strSql =  " SELECT itemid, makerid, itemname, mwdiv, sellyn, regdate, sellstdate, smallimage, favcount "	
 	strSql =  strSql & " from ( "
    strSql =  strSql & " select ROW_NUMBER() OVER (ORDER BY  "&strSqlSort&" ) as RowNum "
    strSql =  strSql & " , t.itemid, i.itemname, i.makerid, i.mwdiv, i.sellyn, i.regdate, i.sellstdate, i.smallimage, t.favcount "
    strSql =  strSql & " from #tmpWish as t "
    strSql =  strSql & " inner join db_analyze_data_raw.dbo.tbl_item as i on i.itemid = t.itemid " 
    strSql =  strSql & ") AS TB " 
	strSql =  strSql &" WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo   
    strSql =  strSql &" order by " & strSqlSort1
    
    rsAnalget.open strSql,dbAnalget,1
		IF Not (rsAnalget.EOF OR rsAnalget.BOF) THEN
			fnGetCleanItemList = rsAnalget.getRows()
		END IF
	rsAnalget.close
	EnD IF	
    End Function
END Class    
%>