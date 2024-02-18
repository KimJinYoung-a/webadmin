<%
'###########################################################
' Description : 온라인 & 오프라인 어드민 장바구니 클래스
' Hieditor : 2011.08.02 한용민 생성
'###########################################################
%>
<script language="JavaScript" src="/js/adminshoppingbag.js"></script>
<%
Class cadminshoppingbag_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fshopname
	public fonlinemwdiv
	public fonlinebuycash
	public fcomm_name
	public fisusing
	public fshopid
	public fitemgubun
	public fitemid
	public fitemoption
	public fregdate
	public flastupdate
	public fmakerid
	public fshopitemname
	public fshopitemoptionname
	public fshopitemprice
	public fshopsuplycash
	public forgsellprice
	public fdiscountsellprice
	public fextbarcode
	public foffimgmain
	public foffimglist
	public foffimgsmall
	public fvatinclude
	public fshopbuyprice
	public fcentermwdiv
	public fcatecdl
	public fcatecdm
	public fcatecdn
	public fonofflinkyn
	public FimageSmall			
	public fchargediv
	public fcomm_cd
	public fdefaultmargin
	public fdefaultsuplymargin
    public fbagidx
    public fuserid
    public fonoffgubun
    public fitemno
	public flogicsipgono
	public flogicsreipgono
	public fbrandipgono
	public fbrandreipgono
	public fsellno
	public fresellno
	public ferrsampleitemno
	public ferrbaditemno
	public ferrrealcheckno
	public fsysstockno
	public frealstockno
	public frequiredStock
	public fsell7days
	public fsell3days
	public fpreorderno
	public fpreordernofix
	public frequire3daystock
	public frequire7daystock
	public frequire14daystock

	''유효재고
    public function getAvailStock()
        getAvailStock = FrealstockNo + Ferrsampleitemno + Ferrbaditemno
    end function
    			    	
	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function	
	
end class

class cadminshoppingbag_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public frectonoffgubun
	public frectuserid
	public Frectshopid
	public FRectisUsing
	public Frectmakerid
	public Frectitemid
	public Frectitemname
	public frectcdl
	public frectcdm
	public frectcds
	public Frectcomm_cd

	'/오프라인 장바구니 리스트	'/common/item/adminshoppingbag.asp
	public sub fadminshoppingbag_off()
		dim sqlStr,i ,sqlsearch
		
		'/정산구분
		if Frectcomm_cd <> "" then
			if Frectcomm_cd = "B088" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B012','B022')"			
			elseif Frectcomm_cd = "B099" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B011','B031')"
			elseif Frectcomm_cd = "B077" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B011','B031','B013')"				
			else
				sqlsearch = sqlsearch & " and d.comm_cd = '"&Frectcomm_cd&"'"
			end if
		end if
		
		'/매장
		if Frectshopid <> "" then
			sqlsearch = sqlsearch & " and b.shopid = '"&Frectshopid&"'"
		end if
		
		'/상품사용여부
		if Frectisusing <> "" then
			sqlsearch = sqlsearch & " and i.isusing = '"&Frectisusing&"'"
		end if
		
		'/브랜드
		if Frectmakerid <> "" then
			sqlsearch = sqlsearch & " and i.makerid = '"&Frectmakerid&"'"
		end if
		
		'/상품코드
		if Frectitemid <> "" then
			sqlsearch = sqlsearch & " and i.shopitemid = "&Frectitemid&""
		end if
		
		'/상품명
		if Frectitemname <> "" then
			sqlsearch = sqlsearch & " and i.shopitemname like '%"&Frectitemname&"%'"
		end if
		
		'/대카테고리
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
        end if
        
        '/중카테고리
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and i.catecdm='" + FRectCDM + "'"
        end if
        
        '/소카테고리
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and i.catecdn='" + FRectCDS + "'"
        end if
		
		'/온라인(ON) & 오프라인 구분(OFF)
		if frectonoffgubun <> "" then
			sqlsearch = sqlsearch + " and b.onoffgubun = '"&frectonoffgubun&"'"
		end if
		
		'/아이디
		if frectuserid <> "" then
			sqlsearch = sqlsearch + " and b.userid = '"&frectuserid&"'"
		end if
		
		'데이터 리스트 
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " b.bagidx ,b.itemno ,b.userid ,b.shopid ,b.onoffgubun"
		sqlStr = sqlStr & " ,i.itemgubun, i.shopitemid, i.itemoption, i.makerid, i.shopitemname, i.shopitemoptionname"
		sqlStr = sqlStr & " ,i.shopitemprice, i.orgsellprice, i.discountsellprice, i.isusing"
		sqlStr = sqlStr & " ,i.extbarcode, i.offimgmain, i.offimglist, i.offimgsmall, i.vatinclude"
		sqlStr = sqlStr & " ,i.centermwdiv, i.catecdl, i.catecdm, i.catecdn,i.onofflinkyn"
		sqlStr = sqlStr & " ,ii.smallimage"
		sqlStr = sqlStr & " ,isnull(s.logicsipgono,0) as logicsipgono, isnull(s.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr & " , isnull(s.brandipgono,0) as brandipgono, isnull(s.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr & " , isnull(s.sellno,0) as sellno, isnull(s.resellno,0) as resellno,isnull(s.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr & " , isnull(s.errbaditemno,0) as errbaditemno, isnull(s.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr & " , isnull(s.sysstockno,0) as sysstockno, isnull(s.realstockno,0) as realstockno, isnull(s.requiredStock,0) as requiredStock"
		sqlStr = sqlStr & " , isnull(s.sell7days,0) as sell7days, isnull(s.sell3days,0) as sell3days ,s.lastupdate"
		sqlStr = sqlStr & " , s.preorderno ,s.preordernofix"
		sqlStr = sqlStr & " ,( (s.sell3days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require3daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require7daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*2) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require14daystock"
		sqlStr = sqlStr & " ,d.chargediv ,d.comm_cd ,d.defaultmargin ,d.defaultsuplymargin"
		sqlStr = sqlStr & " ,(CASE"
		'//온라인이매입이고, 특정이 아니라면 온라인 매입가
		sqlStr = sqlStr & " 	WHEN (i.itemgubun='10' and IsNULL(ii.mwdiv,'')='M' and d.comm_cd not in ('B012','B022')) THEN ii.buycash+IsNULL(o.optaddprice,0)"
		sqlStr = sqlStr & "     WHEN i.shopsuplycash=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopsuplycash END) as shopsuplycash"         ''매입가
		sqlStr = sqlStr & " ,(CASE when i.shopbuyprice=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultsuplymargin,30))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopbuyprice END) as shopbuyprice"           ''매장공급가.
		sqlStr = sqlStr & " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr & " ,IsNULL(ii.mwdiv,'') as mwdiv, (IsNULL(ii.buycash,0)+IsNULL(o.optaddprice,0)) as onlinebuycash, c.comm_name"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_adminshoppingbag b"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " 	on b.itemgubun = i.itemgubun"
		sqlStr = sqlStr & " 	and b.itemid = i.shopitemid"
		sqlStr = sqlStr & " 	and b.itemoption = i.itemoption"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr & " 	on b.shopid = d.shopid"
		sqlStr = sqlStr & "		and i.makerid = d.makerid"
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on b.shopid = u.userid"
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_shopstock_summary s"
		sqlStr = sqlStr & " 	on b.itemgubun = s.itemgubun"
		sqlStr = sqlStr & " 	and b.itemid = s.itemid"
		sqlStr = sqlStr & " 	and b.itemoption = s.itemoption"
		sqlStr = sqlStr & " 	and b.shopid = s.shopid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item ii"
		sqlStr = sqlStr & " 	on i.itemgubun = ii.itemgubun"
		sqlStr = sqlStr & " 	and i.shopitemid = ii.itemid"
		sqlStr = sqlStr & " 	and i.itemgubun = '10'"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_option o"
		sqlStr = sqlStr & "     on i.itemgubun = '10'"
		sqlStr = sqlStr & "     and i.shopitemid = o.itemid"
		sqlStr = sqlStr & "     and i.itemoption = o.itemoption"
		sqlStr = sqlStr & " left join db_jungsan.dbo.tbl_jungsan_comm_code c"
		sqlStr = sqlStr & " 	on d.comm_cd=c.comm_cd"
		sqlStr = sqlStr & " 	and comm_group='Z002' and comm_isdel='N'"
		sqlStr = sqlStr & " where b.itemid <> 0 " & sqlsearch
		'sqlStr = sqlStr & " order by b.shopid asc ,d.comm_cd asc,i.makerid asc ,i.itemgubun asc,i.shopitemid asc ,i.itemoption asc"
		sqlStr = sqlStr & " order by b.bagidx desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)
		
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cadminshoppingbag_item

				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fonlinemwdiv = rsget("mwdiv")
				FItemList(i).fonlinebuycash = rsget("onlinebuycash")
				FItemList(i).fcomm_name = rsget("comm_name")
				FItemList(i).fpreorderno = rsget("preorderno")
				FItemList(i).fpreordernofix = rsget("preordernofix")				
				FItemList(i).frequire3daystock = rsget("require3daystock")
				FItemList(i).frequire7daystock = rsget("require7daystock")
				FItemList(i).frequire14daystock = rsget("require14daystock")
				FItemList(i).fbagidx = rsget("bagidx")
				FItemList(i).fitemno = rsget("itemno")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fonoffgubun = rsget("onoffgubun")				
				FItemList(i).FimageSmall = db2html(rsget("smallimage"))
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemid = rsget("shopitemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fshopitemname = db2html(rsget("shopitemname"))
				FItemList(i).fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).fshopitemprice = rsget("shopitemprice")
				FItemList(i).forgsellprice = rsget("orgsellprice")
				FItemList(i).fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fextbarcode = rsget("extbarcode")
				FItemList(i).foffimgmain = rsget("offimgmain")
				FItemList(i).foffimglist = db2html(rsget("offimglist"))
				FItemList(i).foffimgsmall = db2html(rsget("offimgsmall"))
				FItemList(i).fvatinclude = rsget("vatinclude")
				FItemList(i).fcentermwdiv = rsget("centermwdiv")
				FItemList(i).fcatecdl = rsget("catecdl")
				FItemList(i).fcatecdm = rsget("catecdm")
				FItemList(i).fcatecdn = rsget("catecdn")
				FItemList(i).fonofflinkyn = rsget("onofflinkyn")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fshopsuplycash = rsget("shopsuplycash")
				FItemList(i).fchargediv = rsget("chargediv")
				FItemList(i).fcomm_cd = rsget("comm_cd")
				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")
				FItemList(i).flogicsipgono = rsget("logicsipgono")
				FItemList(i).flogicsreipgono = rsget("logicsreipgono")
				FItemList(i).fbrandipgono = rsget("brandipgono")
				FItemList(i).fbrandreipgono = rsget("brandreipgono")
				FItemList(i).fsellno = rsget("sellno")
				FItemList(i).fresellno = rsget("resellno")
				FItemList(i).ferrsampleitemno = rsget("errsampleitemno")
				FItemList(i).ferrbaditemno = rsget("errbaditemno")
				FItemList(i).ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).fsysstockno = rsget("sysstockno")
				FItemList(i).frealstockno = rsget("realstockno")
				FItemList(i).frequiredStock = rsget("requiredStock")
				FItemList(i).fsell7days = rsget("sell7days")
				FItemList(i).fsell3days = rsget("sell3days")
												
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
											
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/온라인 장바구니 리스트	'/common/item/adminshoppingbag.asp
	public sub fadminshoppingbag_on()

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
end class

'/장바구니 담기 온라인 & 오프라인 공용 함수
function putadminshoppingbag_insert(userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr, reloadsubmit,menupos)
	dim tmpuserid, tmponoffgubun, tmpshopid, tmpitemgubunarr, tmpitemidarr, tmpitemoptionarr, tmpitemnoarr
	dim sql ,i
	
	if userid = "" or onoffgubun = "" or itemgubunarr = "" or itemidarr = "" or itemoptionarr = "" or itemnoarr = "" then exit function
	if shopid = "" then shopid = "NULL"

	tmpitemgubunarr = split(itemgubunarr,",")
	tmpitemidarr = split(itemidarr,",")
	tmpitemoptionarr = split(itemoptionarr,",")
	tmpitemnoarr = split(itemnoarr,",")

	for i = 0 to ubound(tmpitemidarr) - 1
		
		'/장바구니에 같은상품이 존재할경우 기존 상품에 더함 
		sql = ""
		sql = "update b set" + vbcrlf
		sql = sql & " itemno = itemno + "& requestCheckVar(tmpitemnoarr(i),10) &"" + vbcrlf
		sql = sql & " from db_temp.dbo.tbl_adminshoppingbag b" + vbcrlf
		sql = sql & " where itemid = '"& requestCheckVar(tmpitemidarr(i),10) &"'" + vbcrlf
		sql = sql & " and itemoption = '"& requestCheckVar(tmpitemoptionarr(i),4) &"'" + vbcrlf
		sql = sql & " and itemgubun = '"& requestCheckVar(tmpitemgubunarr(i),2) &"'" + vbcrlf
		sql = sql & " and shopid = '"&shopid&"'" + vbcrlf
		sql = sql & " and onoffgubun = '"&onoffgubun&"'" + vbcrlf
		'sql = sql & " and userid = '"&userid&"'" + vbcrlf

		'response.write sql & "<Br>"
		dbget.execute sql		

		sql = ""		
		
		'/장바구니에 같은 상품이 없는 경우 새로등록 온라인 오프라인 구분해서 쿼리해서 넣음
		'//온라인		
		if onoffgubun = "ON" then
		
		'//오프라인				
		elseif onoffgubun = "OFF" then
			sql = "insert into db_temp.dbo.tbl_adminshoppingbag(" + vbcrlf
			sql = sql & " itemgubun ,itemid ,itemoption ,itemno ,userid ,shopid ,onoffgubun)" + vbcrlf
			sql = sql & " 		select i.itemgubun,i.shopitemid,i.itemoption,"& requestCheckVar(tmpitemnoarr(i),10) &"" + vbcrlf
			sql = sql & " 		,'"&userid&"','"&shopid&"','"&onoffgubun&"'" + vbcrlf
			sql = sql & " 		from db_shop.dbo.tbl_shop_item i" + vbcrlf
			sql = sql & " 		left join db_temp.dbo.tbl_adminshoppingbag b" + vbcrlf
			sql = sql & " 			on i.itemgubun = b.itemgubun" + vbcrlf
			sql = sql & " 			and i.shopitemid = b.itemid" + vbcrlf
			sql = sql & " 			and i.itemoption = b.itemoption" + vbcrlf
			'sql = sql & " 			and b.userid = '"&userid&"'" + vbcrlf
			sql = sql & " 			and b.shopid = '"&shopid&"'" + vbcrlf
			sql = sql & " 			and b.onoffgubun = '"&onoffgubun&"'" + vbcrlf
			sql = sql & " 		where b.itemid is null" + vbcrlf
			sql = sql & " 		and i.shopitemid = '"& requestCheckVar(tmpitemidarr(i),10) &"'" + vbcrlf
			sql = sql & " 		and i.itemoption = '"& requestCheckVar(tmpitemoptionarr(i),4) &"'" + vbcrlf
			sql = sql & " 		and i.itemgubun = '"& requestCheckVar(tmpitemgubunarr(i),2) &"'" + vbcrlf

			'response.write sql & "<Br>"
			dbget.execute sql			
		end if
	next
	
	if reloadsubmit <> "" then
		'/장바구니 담은후 페이지 리로드		'/리로드 안할경우 유저가 새로고침 누를 경우 기존 내역에 상품이 계속 합산됨..
	    response.write "<script type='text/javascript'>"
	    response.write 		reloadsubmit & ".location.href='/common/item/adminshoppingbag.asp?onoffgubun="&onoffgubun&"&menupos="&menupos&"';"
	    response.write "</script>"
	    dbget.close() : response.end
	end if
end function
%>