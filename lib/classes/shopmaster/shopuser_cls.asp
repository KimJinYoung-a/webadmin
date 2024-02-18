<%
'###########################################################
' Description : 매장 직원 권한 관련 클래스
' Hieditor : 2011.01.10 한용민 생성
'###########################################################

Class cshopuser_item
	public fid
	public fshopid
	public fregdate
	public ffirstisusing
	public fshopcount
	public fshopfirst
	public fcompany_name
    public firstisusing
    public fshopname
    public fempno
    public fusername
    public fuserid
    public fjoinday
    public fretireday
    public fpart_sn
    public fposit_sn
    public fjob_sn
    public fusermail
    public finterphoneno
    public fextension
    public fdirect070
    public fstatediv
    public fuserimage
    public fusercell
    public fuserphone
    public fmsnmail
	public fjob_name
	public fposit_name
	public fpart_name	
	public fpassword
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class cshopuser_list
    public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public frectid
	public frectempno
	public frectpart_sn
	public frectadminyn
	public frectSearchKey
	public frectSearchString
	
	'//common/offshop/member/shopuser_reg.asp
	public Sub getshopusermember_list()
	    dim sqlStr, i , sqlsearch
	    
	    if frectempno <> "" then
	    	sqlsearch = sqlsearch & " and ut.empno = '"&frectempno&"'"
	    end if
	    
	    sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf 
	    sqlStr = sqlStr + " join db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf
	    sqlStr = sqlStr + " 	on ps.empno = ut.empno" + vbcrlf
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p" + vbcrlf 
		sqlStr = sqlStr + " 	on ut.userid = p.id" + vbcrlf
		sqlStr = sqlStr + " 	and ut.isUsing = 1" & vbcrlf 

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user su" + vbcrlf 
		sqlStr = sqlStr + " on ps.shopid = su.userid" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

	    'response.write sqlStr &"<Br>"	        
	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close
		
		sqlStr = "select top " & CStr(FPageSize*FCurrpage)
		sqlStr = sqlStr & " ut.userid ,ps.shopid ,ps.regdate ,ps.firstisusing ,su.shopname"
		sqlStr = sqlStr & " ,ut.empno , p.password"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf 
	    sqlStr = sqlStr + " join db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf
	    sqlStr = sqlStr + " 	on ps.empno = ut.empno" + vbcrlf
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p" + vbcrlf 
		sqlStr = sqlStr + " 	on ut.userid = p.id" + vbcrlf
		sqlStr = sqlStr + " 	and ut.isUsing = 1" & vbcrlf 

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "		and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user su" + vbcrlf 
		sqlStr = sqlStr + " on ps.shopid = su.userid" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
	    
	    'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new cshopuser_item
				
				FItemList(i).fpassword = rsget("password")
				FItemList(i).fshopname = db2html(rsget("shopname"))
				FItemList(i).fid = rsget("userid")
                FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fregdate = rsget("regdate")
                FItemList(i).firstisusing = rsget("firstisusing")
                FItemList(i).fshopname = db2html(rsget("shopname"))
                FItemList(i).fempno = rsget("empno")
                
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

	'/오프라인팀 직영매장 직원 리스트 '//common/offshop/member/shopuser_list.asp
	public Sub getshopuser_list()
		dim sqlStr,i , sqlsearch
	    
		if frectpart_sn <> "" then
			sqlsearch = sqlsearch & " and ut.part_sn = "&frectpart_sn&""
		end if
		if frectadminyn = "Y" then
			sqlsearch = sqlsearch & " and p.id is not null"
		elseif frectadminyn = "N" then
			sqlsearch = sqlsearch & " and p.id is null"
		end if
		if frectSearchKey = "1" and frectSearchString <> "" then
			sqlsearch = sqlsearch & " and ut.userid ='"&frectSearchString&"'"
		elseif frectSearchKey = "2" and frectSearchString <> "" then
			sqlsearch = sqlsearch & " and ut.username ='"&frectSearchString&"'"
		elseif frectSearchKey = "3" and frectSearchString <> "" then
			sqlsearch = sqlsearch & " and ut.empno ='"&frectSearchString&"'"
		end if
		
	    sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf 
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p" + vbcrlf 
		sqlStr = sqlStr + " 	on ut.userid = p.id" + vbcrlf
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf 
		sqlStr = sqlStr + " 	on ps.empno = ut.empno and ps.firstisusing='Y'" + vbcrlf 
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user su" + vbcrlf 
		sqlStr = sqlStr + " 	on ps.shopid = su.userid" + vbcrlf
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_partInfo as B" + vbcrlf 
		sqlStr = sqlStr + " 	ON ut.part_sn = B.part_sn" + vbcrlf 
		sqlStr = sqlStr + " where ut.isUsing = 1 " & sqlsearch

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf

	    'response.write sqlStr &"<Br>"	        
	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit Sub
			
		sqlStr = "select top " & CStr(FPageSize*FCurrpage)
		sqlStr = sqlStr + " ut.userid ,ut.username ,(ps.shopid) as shopfirst ,su.shopname" + vbcrlf 
		sqlStr = sqlStr + " ,ut.empno ,B.part_name" + vbcrlf
		sqlStr = sqlStr + " ,(select count(*) from db_partner.dbo.tbl_partner_shopuser" + vbcrlf 
		sqlStr = sqlStr + " 	where empno = ut.empno) as shopcount" + vbcrlf 
		sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf 
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p" + vbcrlf 
		sqlStr = sqlStr + " 	on ut.userid = p.id" + vbcrlf
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf 
		sqlStr = sqlStr + " 	on ps.empno = ut.empno and ps.firstisusing='Y'" + vbcrlf 
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user su" + vbcrlf 
		sqlStr = sqlStr + " 	on ps.shopid = su.userid" + vbcrlf
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_partInfo as B" + vbcrlf 
		sqlStr = sqlStr + " 	ON ut.part_sn = B.part_sn" + vbcrlf 
		sqlStr = sqlStr + " where ut.isUsing = 1 " & sqlsearch

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr + " order by ut.part_sn desc ,ut.empno asc"
	    
	    'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new cshopuser_item
				
					FItemList(i).fpart_name       = db2html(rsget("part_name"))
					FItemList(i).fempno       = rsget("empno")
					FItemList(i).fshopname = db2html(rsget("shopname"))				
					FItemList(i).fid       = rsget("userid")
					FItemList(i).fcompany_name       = db2html(rsget("username"))
					FItemList(i).fshopcount       = rsget("shopcount")
					FItemList(i).fshopfirst       = rsget("shopfirst")
                
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

	'/오프라인팀 사무실 직원 리스트 '/common/offshop/member/offlinemember_list.asp
	public Sub getofflinemember_list()
		dim sqlStr,i
	
		sqlStr = "select top 500"
		sqlStr = sqlStr + " t.empno, t.username, t.userid, t.joinday, t.retireday, t.part_sn, t.posit_sn"
		sqlStr = sqlStr + " ,t.job_sn, t.usermail ,t.interphoneno, t.extension, t.direct070, t.jobdetail"
		sqlStr = sqlStr + " ,t.statediv, t.userimage , t.usercell, t.userphone, t.msnmail"
		sqlStr = sqlStr + " ,D.job_name, C.posit_name, B.part_name"
		sqlStr = sqlStr + " ,su.shopname,(ps.shopid) as shopfirst"
		sqlStr = sqlStr + " ,(select count(*) from db_partner.dbo.tbl_partner_shopuser where empno = t.empno) as shopcount"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten t"
		sqlStr = sqlStr + " join db_partner.dbo.tbl_partner p"
		sqlStr = sqlStr + " 	on t.userid = p.id"
		sqlStr = sqlStr + " 	and p.isusing='Y' and t.isusing=1"
		sqlStr = sqlStr + " 	and t.part_sn in (13,18)"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_partInfo as B"
		sqlStr = sqlStr + " 	ON t.part_sn = B.part_sn"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_positInfo as C"
		sqlStr = sqlStr + " 	ON t.posit_sn = C.posit_sn"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_JobInfo as D"
		sqlStr = sqlStr + " 	ON t.job_sn = D.job_sn"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner_shopuser ps"
		sqlStr = sqlStr + " 	on ps.empno = t.empno and ps.firstisusing='Y'"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user su"
		sqlStr = sqlStr + " 	on ps.shopid = su.userid"
		sqlStr = sqlStr + " order by t.part_sn asc, t.posit_sn asc"
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount

		redim preserve FItemList(FTotalCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cshopuser_item

					FItemList(i).fshopname       = rsget("shopname")
					FItemList(i).fshopfirst       = rsget("shopfirst")
					FItemList(i).fshopcount       = rsget("shopcount")
					FItemList(i).fempno       = rsget("empno")
					FItemList(i).fusername       = db2html(rsget("username"))
					FItemList(i).fuserid       = rsget("userid")
					FItemList(i).fjoinday       = rsget("joinday")
					FItemList(i).fretireday       = rsget("retireday")
					FItemList(i).fpart_sn       = rsget("part_sn")
					FItemList(i).fposit_sn       = rsget("posit_sn")
					FItemList(i).fjob_sn       = rsget("job_sn")
					FItemList(i).fusermail       = db2html(rsget("usermail"))
					FItemList(i).finterphoneno       = rsget("interphoneno")
					FItemList(i).fextension       = rsget("extension")
					FItemList(i).fdirect070       = rsget("direct070")
					FItemList(i).fstatediv       = rsget("statediv")																																																												
					FItemList(i).fuserimage       = db2html(rsget("userimage"))
					FItemList(i).fusercell       = rsget("usercell")
					FItemList(i).fuserphone       = rsget("userphone")
					FItemList(i).fmsnmail       = db2html(rsget("msnmail"))
					FItemList(i).fjob_name       = rsget("job_name")
					FItemList(i).fposit_name       = rsget("posit_name")
					FItemList(i).fpart_name       = rsget("part_name")
																									
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub
	
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
%>