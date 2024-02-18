<%
'####################################################
' Description :  오프라인 결제 정산 클래스
' History : 2013.10.24 한용민 생성
'####################################################

class Cpayment_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fyyyymmdd
	public fdpart
	public Fcashsum
	public Fcashcnt
	public fidx
	public fshopid
	public fcnt100000won
	public fcnt50000won
	public fcnt10000won
	public fcnt5000won
	public fcnt1000won
	public fcnt500won
	public fcnt100won
	public fcnt50won
	public fcnt10won
	public fvaultcash
	public fjungsanadminid
	public fdepositadminid
	public fisusing
	public fcodeid
	public fcodename
	public fdetailidx
	public fmasteridx
	public fetcwon
	public fposid
	public fposidcnt
	public fbigo
end class

Class Cpayment
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public ftmpyyyymmdd
	public ftmpposidcnt
	public fposidarr
	public frectdatefg
	public FRectStartDay
	public FRectEndDay
	public FRectShopID
	public frectmasteridx
	public FRectInc3pl
	
	'//admin/offshop/payment/cash_management.asp
	public Sub Getcash_management()
		Dim sql,i
		
		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql + " isnull(t.cashsum,0) as cashsum, isnull(t.cashcnt,0) as cashcnt, t.yyyymmdd, t.dpart, t.shopid, t.posid"
		sql = sql + " , c.idx, isnull(c.cnt100000won,0) as cnt100000won, isnull(c.cnt50000won,0) as cnt50000won, isnull(c.cnt10000won,0) as cnt10000won"
		sql = sql + " , isnull(c.cnt5000won,0) as cnt5000won, isnull(c.cnt1000won,0) as cnt1000won, isnull(c.cnt500won,0) as cnt500won"
		sql = sql + " , isnull(c.cnt100won,0) as cnt100won, isnull(c.cnt50won,0) as cnt50won, isnull(c.cnt10won,0) as cnt10won"
		sql = sql + " , isnull(c.vaultcash,0) as vaultcash, c.jungsanadminid, c.depositadminid, c.bigo"
		sql = sql + " from ("
		sql = sql + " 	select"
		sql = sql + " 	substring(orderno,10,2) as posid"
		sql = sql + " 	,m.shopid"
		sql = sql + " 	,sum(cashsum) as 'cashsum'"
		sql = sql + " 	,sum(case when jumunmethod='01' then 1 else 0 end) as 'cashcnt'"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql + " 	,convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart"
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql + " 	,(m.IXyyyymmdd) as yyyymmdd, datepart(w,m.IXyyyymmdd) as dpart"
		else
			sql = sql + " 	,convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart"
		end if
		
		sql = sql + " 	from [db_shop].[dbo].tbl_shopjumun_master m"
		sql = sql + " left join db_partner.dbo.tbl_partner p"
	    sql = sql + "       on m.shopid=p.id "		
		sql = sql + " 	where m.cancelyn='N'"

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sql = sql & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sql = sql & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopID<>"" then
			sql = sql + " 	and m.shopid='" + FRectShopID + "'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sql = sql + " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sql = sql + " 	and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " 	and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sql = sql + " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		sql = sql + " 	group by substring(orderno,10,2), m.shopid"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql + " 	,convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)"
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql + " 	,m.IXyyyymmdd"
		else
			sql = sql + " 	,convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)"
		end if
		sql = sql + " ) as t"
		sql = sql + " left join db_shop.dbo.tbl_shop_cash_management c"
		sql = sql + " 	on t.yyyymmdd = c.yyyymmdd"
		sql = sql + " 	and t.shopid=c.shopid"
		sql = sql + " 	and c.isusing='Y'"
		sql = sql + " 	and t.posid=c.posid"
		
		if FRectShopID<>"" then
			sql = sql + " 	and c.shopid='" + FRectShopID + "'"
		end if
		
		sql = sql + " order by t.yyyymmdd desc, t.posid asc"

		'response.write sql & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sql,dbget,1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		ftotalcount = rsget.RecordCount
			
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cpayment_oneitem
				
				'/셀병합 처리를 위해 포스가 2대이상일경우 배열에 값을 저장
				if rsget("yyyymmdd")=ftmpyyyymmdd then
					ftmpposidcnt = ftmpposidcnt + 1
				else
					ftmpposidcnt = "1"
				end if
				if ftmpposidcnt>1 then
					if instr(fposidarr,rsget("yyyymmdd"))="0" then
						fposidarr=fposidarr & rsget("yyyymmdd") & "|" & ftmpposidcnt & ","
					else
						fposidarr = left(fposidarr, instr(fposidarr,rsget("yyyymmdd"))-1)
						fposidarr=fposidarr & rsget("yyyymmdd") & "|" & ftmpposidcnt & ","
					end if
				end if
				ftmpyyyymmdd = rsget("yyyymmdd")
				FItemList(i).fposidcnt = ftmpposidcnt
				
				FItemList(i).fbigo = db2html(rsget("bigo"))
				FItemList(i).fposid = rsget("posid")
			    FItemList(i).fyyyymmdd = rsget("yyyymmdd")
			    FItemList(i).fcashcnt           = rsget("cashcnt")
				FItemList(i).fcashsum           = rsget("cashsum")
				FItemList(i).fdpart           = rsget("dpart")
				FItemList(i).fshopid           = rsget("shopid")
				FItemList(i).fidx           = rsget("idx")
				FItemList(i).fcnt100000won           = rsget("cnt100000won")
				FItemList(i).fcnt50000won           = rsget("cnt50000won")
				FItemList(i).fcnt10000won           = rsget("cnt10000won")
				FItemList(i).fcnt5000won           = rsget("cnt5000won")
				FItemList(i).fcnt1000won           = rsget("cnt1000won")
				FItemList(i).fcnt500won           = rsget("cnt500won")
				FItemList(i).fcnt100won           = rsget("cnt100won")
				FItemList(i).fcnt50won           = rsget("cnt50won")
				FItemList(i).fcnt10won           = rsget("cnt10won")
				FItemList(i).fvaultcash           = rsget("vaultcash")
				FItemList(i).fjungsanadminid           = db2html(rsget("jungsanadminid"))
				FItemList(i).fdepositadminid           = db2html(rsget("depositadminid"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/payment/cash_management.asp
	public Sub Getcash_management_etc()
		Dim sql,i

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql + " oc.codeid, oc.codename, e.detailidx, e.masteridx, isnull(e.etcwon,0) as etcwon"
		sql = sql + " from db_shop.dbo.tbl_offshop_commoncode oc"
		sql = sql + " left join db_shop.dbo.tbl_shop_cash_management_etc e"
		sql = sql + " 	on oc.codeid = e.etctype"
		sql = sql + " 	and e.isusing='Y'"
		
		if frectmasteridx <> "" then
			sql = sql & " and e.masteridx="& frectmasteridx &""
		else
			sql = sql & " and e.masteridx=0"
		end if			
	
		sql = sql + " where oc.codekind='etctype'"
		sql = sql + " and oc.codegroup='MAIN'"
		sql = sql + " and oc.useyn='Y'"
		sql = sql + " order by oc.orderno asc"

		'response.write sql & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sql,dbget,1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		ftotalcount = rsget.RecordCount
			
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cpayment_oneitem

			    FItemList(i).fcodeid = rsget("codeid")
			    FItemList(i).fcodename = db2html(rsget("codename"))
			    FItemList(i).fdetailidx = rsget("detailidx")
			    FItemList(i).fmasteridx = rsget("masteridx")
			    FItemList(i).fetcwon = rsget("etcwon")

				i=i+1
				rsget.moveNext
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

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		ftmpposidcnt=1
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class
%>