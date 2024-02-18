<%
'###########################################################
' Description :  옥션 상품 관리 클래스
' History : 2007.09.11 한용민 생성
'###########################################################
CLASS CCategoryPrdItem

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FDiscountRate

	public FCodeLarge
	public FCodeMid
	public FCodeSmall

	public Fidx
	
	public FItemID
	public FItemName
	public Fitemcontents
	public FSellcash
	public FSellYn
	public FDispYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public FItemSource
	public FSourceArea
	public FImageAddStr
	public FImageInfoStr
	
	
	public FImageSmall
	public FImageList
	public FImageList120	
	public FImageBasic
	Public FImageIcon1
	public FImageBasicIcon
	
	public FAddimageGubun
	public FAddimageSmall
	public FAddimage
	
	public FMakerID
	public Fitemcontent
	public FRegdate
	
	public Fimgstory
	public Fdesignercomment
	public Fitemgubun
	public FPoints
	
	public FDeliverytype

	public Fevalcnt
	public Ffreeprizeyn
	public Fsatisfyitemyn
	public Fitemcouponyn
	public Flimitsoldoutyn
	public Fcontents

	public Fdesignerid
	public Fcd1
	public Fsellsum
	public Fsoccomment
	public Fsoclogo

	public FSaleYn
	public FOrgPrice
	public FSailPrice
	public FEventPrice
	public FImageMain
	public FlinkURL

	public FSpecialuseritem

	public FEvalComments

	public Fcdlarge
	public Fcdmid
	public Fnmmid
	
	
	Public FItemSize
	public FOrderComment
	public FImageAddContentStr
	public FMakerName
	public FUsingHTML
	public FMileage
	public Ftodaydeliver
	public Fdeliverarea
	public FReipgodate
	public FIsMobileItem
	public FFingerId
	public FOptionCnt
	public FItemCouponType
	public FItemCouponValue
	public FReipgoItemYN
	public FItemDiv
	public Fcurritemcouponidx
	
	public Fsocname_kor
	public FSpecialbrand
	public Fsocname
	public Fdgncomment

	public Fstreetusing
	public Fisusing
	public Fuserdiv
	public FNewitem
		
end CLASS

class Cauctionitem
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub

	public idx
	public ten_itemid
	public ten_option
	public auction_cate_code
	public auction_cadal
	public ten_makerid
	public ten_itemname
	public item_stats
	public panmae_area
	public wonsanji
	public auction_realsel
	public auction_telsel
	public auction_medic
	public auction_gungang
	public auction_sik
	public auction_level
	public ten_jaego
	public auction_div_type
	public ten_itemcontent
	public auction_isusing
	public ten_jaego_isusing
	public smallimage
	public foptioncnt
	
	public Fsellyn
	public Fdispyn
	public Flimityn
	public FLimitNo
	public FLimitSold
	public Fdanjongyn
	public fsellcash
	public fbuycash
	
	public FImageMain
	public FImageList
	public FImageList120
	public FImageSmall
	public FImageBasic
	public FImageBasicIcon
	public FImageInfoStr
	
	public function IsSoldOut()		'품절인지 아닌지
		IsSoldOut = FSellYn<>"Y" or ((FLimitYn<>"N") and (FLimitNo-FLimitSold<1))
	end function
	
	public function GetCalcuMarginRate	'마진율 계산
		GetCalcuMarginRate = 0
		if fsellcash<>0 then
			GetCalcuMarginRate = 100-CLng(fbuycash/fsellcash*100*100)/100
		end if
	end function
end class
'##################################################################
class Cfitem					'재고용 클래스
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx				'인덱스번호
	public fitemgubun		'상품구분
	public fitemid			'상품번호
	public fitemoption		'옵션코드	
	public fitemname		'상품명
	public fitemoptionname	'옵션명
	public fmakerid			'브랜드id
	public fregdate			'등록일
	public freguserid		'지시자id	
	public forderingdate	'작업지시일
	public fauctionusername	'재고파악한사람
	public fauctionstartdate	'재고파악일시
	public fbasicstock		'재고파악재고
	public frealstock		'재고파악 실사갯수
	public ferrstock		'오차
	public ffinishuserid	'완료자id
	public fstatecd			'상태코드
	public deleteyn			'삭제여부
	public makerid			'검색에필요한브랜드id
	public fstats			'상태
	public fbigo			'비고
	public foptioncnt		'옵션비교시 필요한 변수
	public fitemcontent
	public FImageSmall
	
end class
'##################################################################
class Cauctionlist
	public flist
	
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	
	public Frectidx					'인덱스 값을 받기 위한 변수
	public Frectitemid				'상품id 값을 받기 위한 변수
	public frectmakerid				'브랜드 값을 받기 위한 변수
	public frectauction
	public frectten
	public frectmagin
	public fauction_category
	public frectevt_code

	'/admin/auctionadd_event.asp
	public Sub feventitem_list()
		dim sql ,i 

			sql = "select"
			sql = sql & " a.itemid , a.makerid , a.itemname , a.smallimage"
			sql = sql & " ,b.optionname"
			sql = sql & " from db_item.dbo.tbl_item a"
			sql = sql & " left join db_item.dbo.tbl_item_option b"
			sql = sql & " on a.itemid = b.itemid "
			sql = sql & " left join db_event.dbo.tbl_eventitem c"
			sql = sql & " on a.itemid = c.itemid"
			sql = sql & " where 1=1"
			sql = sql & " and sellyn = 'Y'"
			
			if frectevt_code <> "" then
				sql = sql & " and c.evt_code = "& frectevt_code&""
			else 
				sql = sql & " and c.evt_code = 00000"			
			end if
			
			'response.write sql
			rsget.open sql,dbget,1
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem
					
					flist(i).fitemid = rsget("itemid")
					flist(i).fmakerid = rsget("makerid")
					flist(i).fitemname = db2html(rsget("itemname"))
					flist(i).fitemoptionname = rsget("optionname")	
					flist(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")				
					
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub
	
	'//admin/auction/auctionlist.asp 
	public sub fauctionlist			
		dim sql , i ,sqlcount
	
			sqlcount = "select count(a.idx) as cnt"
			sqlcount = sqlcount & " from [db_item].dbo.tbl_auction a"
			sqlcount = sqlcount & " left join [db_item].[dbo].tbl_item_option b"
			sqlcount = sqlcount & " on a.ten_itemid = b.itemid and a.ten_option = b.itemoption"
			sqlcount = sqlcount & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
			sqlcount = sqlcount & " on a.ten_itemid = c.itemid and a.ten_option = isnull(c.itemoptionname,'')"
			sqlcount = sqlcount & " left join [db_item].[dbo].tbl_item d"
			sqlcount = sqlcount & " on a.ten_itemid = d.itemid"
			sqlcount = sqlcount & " left join [db_item].[dbo].tbl_item_contents f"
			sqlcount = sqlcount & " on a.ten_itemid = f.itemid"
			sqlcount = sqlcount & " where 1=1" 
			
		if frectauction <> "" then
			sql = sql & " and a.auction_isusing = '"& frectauction &"'"
		end if
		if frectmakerid <> "" then
			sql = sql & " and d.makerid = '"& frectmakerid &"'"
		end if
		
		if fauction_category <> "" then
			sql = sql & " and a.auction_cate_code = '"& fauction_category &"'"
		end if
		if frectten <> "" then
			if frectten = "y" then
				sql = sql & " and c.realstock >= '10'"
			else 
				sql = sql & " and realstock < '10'"
			end if	
		end if
		if frectmagin <> "" then
			if frectmagin = "20" then
				sql = sql & " and (100-(d.buycash/d.sellcash*100*100)/100) >= 20"
			else
				sql = sql & " and (100-(d.buycash/d.sellcash*100*100)/100) < 20"
			end if	
		end if
		
		'response.write sqlcount&"<br>"
		rsget.open sqlcount,dbget,1
		FTotalCount = rsget("cnt")				'총레코드 수에 인덱스카운트를 넣고
		rsget.close
		
			sql = "select top "& FPageSize*FCurrpage&"" 
			sql = sql & " a.idx,a.ten_itemid,a.ten_option,a.auction_realsel,a.auction_isusing"
			sql = sql & " , isnull(c.realstock,'0') as realstock"
			sql = sql & " , d.makerid,d.itemname,f.itemcontent,d.SellYn,d.LimitYn,d.LimitNo,d.LimitSold"
			sql = sql & " ,d.danjongyn,d.sellcash,d.buycash ,d.mainimage,d.listimage,d.basicimage"
			sql = sql & " ,d.smallimage"
			sql = sql & " ,a.auction_cate_code"
			sql = sql & " from [db_item].dbo.tbl_auction a"
			sql = sql & " left join [db_item].[dbo].tbl_item_option b"
			sql = sql & " on a.ten_itemid = b.itemid and a.ten_option = b.itemoption"
			sql = sql & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
			sql = sql & " on a.ten_itemid = c.itemid and a.ten_option = isnull(c.itemoptionname,'')"
			sql = sql & " left join [db_item].[dbo].tbl_item d"
			sql = sql & " on a.ten_itemid = d.itemid"
			sql = sql & " left join [db_item].[dbo].tbl_item_contents f"
			sql = sql & " on a.ten_itemid = f.itemid"
			sql = sql & " where 1=1" 
		
		if frectauction <> "" then
			sql = sql & " and a.auction_isusing = '"& frectauction &"'"
		end if
		if frectmakerid <> "" then
			sql = sql & " and d.makerid = '"& frectmakerid &"'"
		end if
		
		if fauction_category <> "" then
			sql = sql & " and a.auction_cate_code = '"& fauction_category &"'"
		end if
		if frectten <> "" then
			if frectten = "y" then
				sql = sql & " and c.realstock >= '10'"
			else 
				sql = sql & " and realstock < '10'"
			end if	
		end if
		if frectmagin <> "" then
			if frectmagin = "20" then
				sql = sql & " and (100-(d.buycash/d.sellcash*100*100)/100) >= 20"
			else
				sql = sql & " and (100-(d.buycash/d.sellcash*100*100)/100) < 20"
			end if	
		end if
		
		
		sql = sql & " order by a.regdate  desc" 
		'response.write sql&"<br>"
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		
		FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1	
		redim flist(FResultCount)
		i = 0
		
			if not rsget.eof then				'레코드의 첫번째가 아니라면
				rsget.absolutepage = FCurrPage
				do until rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cauctionitem 			'클래스를 넣고
						
						flist(i).idx = rsget("idx")
						flist(i).ten_itemid = rsget("ten_itemid")
						flist(i).ten_option = rsget("ten_option")
						flist(i).auction_cate_code = rsget("auction_cate_code")
						flist(i).ten_makerid = rsget("makerid")
						flist(i).ten_itemname = rsget("itemname")
						flist(i).auction_realsel = rsget("auction_realsel")
						flist(i).ten_jaego = rsget("realstock")
						flist(i).ten_itemcontent = db2html(rsget("itemcontent"))
						flist(i).auction_isusing = rsget("auction_isusing")
						flist(i).Fsellyn = rsget("sellyn")
						flist(i).Flimityn = rsget("limityn")
						flist(i).FLimitNo = rsget("LimitNo")
						flist(i).FLimitSold = rsget("LimitSold")
						flist(i).Fdanjongyn = rsget("danjongyn")
						flist(i).fsellcash = rsget("sellcash")
						flist(i).fbuycash = rsget("buycash")
						flist(i).FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("mainimage")
						flist(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("listimage")
						flist(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("smallimage")
						flist(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("basicimage")									
						rsget.movenext
					i = i+1
					
				loop
			end if
		rsget.close
	end Sub

	'//admin/auction/auctionedit.asp 
	public sub fauctionedit		
		dim sql 
		
		sql = "select" 
		sql = sql & " a.idx,a.ten_itemid,a.ten_option,a.auction_realsel,a.auction_isusing"
		sql = sql & " , isnull(c.realstock,'0') as realstock"
		sql = sql & " , d.makerid,d.itemname,f.itemcontent,d.SellYn,d.LimitYn,d.LimitNo,d.LimitSold"
		sql = sql & " ,d.danjongyn,d.sellcash,d.buycash ,d.mainimage,d.listimage,d.basicimage"
		sql = sql & " ,d.smallimage"
		sql = sql & " ,a.auction_cate_code"
		sql = sql & " from [db_item].dbo.tbl_auction a"
		sql = sql & " left join [db_item].[dbo].tbl_item_option b"
		sql = sql & " on a.ten_itemid = b.itemid and a.ten_option = b.itemoption"
		sql = sql & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
		sql = sql & " on a.ten_itemid = c.itemid and a.ten_option = isnull(c.itemoptionname,'')"
		sql = sql & " left join [db_item].[dbo].tbl_item d"
		sql = sql & " on a.ten_itemid = d.itemid"
		sql = sql & " left join [db_item].[dbo].tbl_item_contents f"
		sql = sql & " on a.ten_itemid = f.itemid"
		sql = sql & " where 1=1" 
		
		if frectidx <> "" then
			sql = sql & "and a.idx = "& frectidx &""
		end if
		
		'response.write sql&"<br>"
		rsget.open sql,dbget,1
		ftotalcount = rsget.recordcount
		
			if not rsget.eof then
				set flist(i) = new Cauctionitem
						
				flist(i).idx = rsget("idx")
				flist(i).ten_itemid = rsget("ten_itemid")
				flist(i).ten_option = rsget("ten_option")
				flist(i).auction_cate_code = rsget("auction_cate_code")
				flist(i).ten_makerid = rsget("makerid")
				flist(i).ten_itemname = rsget("itemname")
				flist(i).auction_realsel = rsget("auction_realsel")
				flist(i).ten_jaego = rsget("realstock")
				flist(i).ten_itemcontent = db2html(rsget("itemcontent"))
				flist(i).auction_isusing = rsget("auction_isusing")
				flist(i).Fsellyn = rsget("sellyn")
				flist(i).Flimityn = rsget("limityn")
				flist(i).FLimitNo = rsget("LimitNo")
				flist(i).FLimitSold = rsget("LimitSold")
				flist(i).Fdanjongyn = rsget("danjongyn")
				flist(i).fsellcash = rsget("sellcash")
				flist(i).fbuycash = rsget("buycash")
				flist(i).FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("mainimage")
				flist(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("listimage")
				flist(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("smallimage")
				flist(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("basicimage")
			end if
		rsget.close
	end Sub

	''상품 옵션별로 대량 재고 검색  '//admin/auction/auction_process
	public Sub fwritelist_daerang()
		dim sql554 ,i 

			sql554 = "select f.itemcontent, a.itemid, a.makerid ," 
			sql554 = sql554 & " isnull(b.itemoption,'0000') as itemoption,"
			sql554 = sql554 & " a.itemname, b.optionname,"
			sql554 = sql554 & " isnull(c.realstock,'0') as realstock"
			sql554 = sql554 & " from [db_item].[dbo].tbl_item a"
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_option b"
			sql554 = sql554 & " on a.itemid = b.itemid"
			sql554 = sql554 & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
			sql554 = sql554 & " on a.itemid = c.itemid and b.optionname = isnull(c.itemoptionname,'')"
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_contents f"
			sql554 = sql554 & " on a.itemid = f.itemid"	
			sql554 = sql554 & " where a.itemid in ("&Frectitemid&")"

			'response.write sql554
			rsget.open sql554,dbget,1			
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem							'클래스넣고
					
					flist(i).fitemid = rsget("itemid")		'상품옵션이름넣고	
					flist(i).fitemoptionname = rsget("optionname")		'상품옵션이름넣고	
					flist(i).fitemoption = rsget("itemoption")			'상품옵션코드넣고
					flist(i).frealstock = rsget("realstock")			'텐재고
					flist(i).fitemcontent = db2html(rsget("itemcontent"))			'상품설명
					
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub
'##################################################################	
	public Sub fwritelist_group()				'상품 옵션별로 그룹으로 묶어서 상품코드 추출
		dim sql554 ,i 

			sql554 = "select itemid" 
			sql554 = sql554 & " from [db_item].[dbo].tbl_item"
			sql554 = sql554 & " where itemid in ("&Frectitemid&")"
			sql554 = sql554 & " group by itemid"
			'response.write sql554
			rsget.open sql554,dbget,1			
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem
					
					flist(i).fitemid = rsget("itemid")
					
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub			

	'// /admin/auction/auctionadd.asp
	public Sub fwritelist()				'상품 옵션별로 재고 검색
		dim sql554 ,i 

			sql554 = "select f.itemcontent, a.itemid, a.makerid ," 
			sql554 = sql554 & " isnull(b.itemoption,'0000') as itemoption,"
			sql554 = sql554 & " a.itemname, b.optionname,"
			sql554 = sql554 & " isnull(c.realstock,'0') as realstock"
			sql554 = sql554 & " from [db_item].[dbo].tbl_item a"
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_option b"
			sql554 = sql554 & " on a.itemid = b.itemid"
			sql554 = sql554 & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
			sql554 = sql554 & " on a.itemid = c.itemid and b.optionname = isnull(c.itemoptionname,'')"
			sql554 = sql554 & " left join [db_item].[dbo].tbl_item_contents f"
			sql554 = sql554 & " on a.itemid = f.itemid"	
			sql554 = sql554 & " where a.itemid = '"& Frectitemid &"'"
			
			'response.write sql554
			rsget.open sql554,dbget,1			
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cfitem							'클래스넣고
					
					flist(i).fitemoptionname = rsget("optionname")		'상품옵션이름넣고	
					flist(i).fitemoption = rsget("itemoption")			'상품옵션코드넣고
					flist(i).frealstock = rsget("realstock")			'텐재고
					flist(i).fitemcontent = db2html(rsget("itemcontent"))			'상품설명
					
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub

'##################################################################
	public Sub fwritelist_auction()				'상품 옵션별로 재고 검색후 저장을 위한 클래스
		dim sql550 ,i 
			
			sql550 = "select a.idx,a.ten_itemid,a.ten_option,a.auction_cate_code"
			sql550 = sql550 & " ,a.auction_realsel,a.auction_isusing,"
			sql550 = sql550 & " isnull(c.realstock,'0') as realstock,"
			sql550 = sql550 & " d.makerid,d.itemname,f.itemcontent"
			sql550 = sql550 & " ,d.mainimage,d.listimage,d.basicimage,d.smallimage"
			sql550 = sql550 & " from [db_item].dbo.tbl_auction a"
			sql550 = sql550 & " left join [db_item].[dbo].tbl_item_option b"
			sql550 = sql550 & " on a.ten_itemid = b.itemid and a.ten_option = b.itemoption"
			sql550 = sql550 & " left join [db_summary].dbo.tbl_current_logisstock_summary c"
			sql550 = sql550 & " on a.ten_itemid = c.itemid and a.ten_option = isnull(c.itemoptionname,'')"
			sql550 = sql550 & " left join [db_item].[dbo].tbl_item d"
			sql550 = sql550 & " on a.ten_itemid = d.itemid"
			sql550 = sql550 & " left join [db_item].[dbo].tbl_item_contents f"
			sql550 = sql550 & " on a.ten_itemid = f.itemid"		
			sql550 = sql550 & " where 1=1 and a.ten_itemid = '"& Frectitemid &"'"
			
			'response.write sql550&"<br>"
			rsget.open sql550,dbget,1			
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cauctionitem							'클래스넣고
					
						flist(i).idx = rsget("idx")
						flist(i).ten_itemid = rsget("ten_itemid")
						flist(i).ten_option = rsget("ten_option")
						flist(i).auction_cate_code = rsget("auction_cate_code")
						flist(i).ten_makerid = rsget("makerid")
						flist(i).ten_itemname = rsget("itemname")
						flist(i).auction_realsel = rsget("auction_realsel")
						flist(i).ten_jaego = rsget("realstock")
						flist(i).ten_itemcontent = db2html(rsget("itemcontent"))
						flist(i).auction_isusing = rsget("auction_isusing")
						flist(i).FImageMain = rsget("mainimage")
						flist(i).FImageList = rsget("listimage")
						flist(i).FImageBasic = rsget("basicimage")
						flist(i).FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("mainimage")
						flist(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("listimage")
						flist(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("smallimage")
						flist(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("ten_itemid")) + "/" + rsget("basicimage")
										
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub

'//admin/auction/auction_xml_new.asp 옥션 솔루션에 맞게 엑셀이나 xml 파일로 출력하기 위한 클래스
public Sub fauction_excel()
	dim sql ,i 

	sql = "select"
		sql = sql & " a.itemid,a.makerid,a.itemname,a.SellYn,a.LimitYn,a.LimitNo"
		sql = sql & " ,a.LimitSold ,a.danjongyn,a.sellcash,a.buycash ,a.mainimage"
		sql = sql & " ,a.listimage,a.basicimage ,a.smallimage ,a.optioncnt"
		sql = sql & " ,b.itemcontent"
		sql = sql & " ,c.auction_cate_code"
		sql = sql & " from [db_item].[dbo].tbl_item a"
		sql = sql & " join [db_item].[dbo].tbl_item_contents b" 
		sql = sql & " on a.itemid = b.itemid" 
		sql = sql & " join ("
		sql = sql & " 	select ten_itemid , auction_cate_code"
		sql = sql & " 	from [db_item].dbo.tbl_auction"
		sql = sql & " 	where 1=1"
			if frectitemid <> "" then
				sql = sql & " and ten_itemid in ("& frectitemid &")" 
			end if
		sql = sql & " 	group by ten_itemid , auction_cate_code"
		sql = sql & " 	) as c"
		sql = sql & " on a.itemid = c.ten_itemid"

	'response.write sql&"<br>"
	rsget.open sql,dbget,1			
	
	FTotalCount = rsget.recordcount
   	redim flist(FTotalCount)
   	i = 0
   	
   	if not rsget.EOF  then
		do until rsget.eof

			set flist(i) = new Cauctionitem							'클래스넣고

			flist(i).foptioncnt = rsget("optioncnt")	
			flist(i).ten_itemid = rsget("itemid")
			flist(i).auction_cate_code = rsget("auction_cate_code")
			flist(i).ten_makerid = db2html(rsget("makerid"))
			flist(i).ten_itemname = db2html(rsget("itemname"))
			flist(i).ten_itemcontent = db2html(rsget("itemcontent"))
			flist(i).fsellcash = rsget("sellcash")
			flist(i).FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
			flist(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
			flist(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")
			flist(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage")
			
			rsget.moveNext		   
			i=i+1
		loop
	end if
	rsget.close
end sub	
	
'//admin/auction/auction_xml_new.asp 옥션 솔루션에 맞게 엑셀이나 xml 파일로 출력하기 위한 클래스( 상세이미지)
public Sub fauction_excel_infoimage()				
	dim sql550 ,t
	
		sql550 = "select ITEMID, ADDIMAGE_400"
		sql550 = sql550 & " from db_item.dbo.tbl_item_addimage"
		sql550 = sql550 & " where 1=1 and imgtype=1 and itemid = "&frectitemid&""
	
		'response.write sql550&"<br>"
		rsget.open sql550,dbget,1			
		
		FTotalCount = rsget.recordcount
	   	redim flist(FTotalCount)
	   	t = 0
	   	
	   	if not rsget.EOF  then
			do until rsget.eof
				set flist(t) = new Cauctionitem							'클래스넣고
				
					flist(t).FImageInfoStr = "http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(rsget("ITEMID")) + "/" + rsget("ADDIMAGE_400")

				rsget.moveNext		   
				t=t+1
			loop
		end if
		rsget.close
	end sub		

'//admin/auction/auction_xml_new.asp
public Sub fitemid_output()				'인덱스명을 가지고 옥션테이블에 상품명을 추출해 낸다.
		dim sql550 ,i 
			
			sql550 = "select" 
			sql550 = sql550 & " ten_itemid"
			sql550 = sql550 & " from [db_item].[dbo].tbl_auction"
			sql550 = sql550 & " where 1=1 and idx in("&Frectidx&")"
			sql550 = sql550 & " group by ten_itemid"

			'response.write sql550
			rsget.open sql550,dbget,1			
			
			FTotalCount = rsget.recordcount
		   	redim flist(FTotalCount)
		   	i = 0
		   	
		   	if not rsget.EOF  then
				do until rsget.eof
					set flist(i) = new Cauctionitem							'클래스넣고
					
						flist(i).ten_itemid = rsget("ten_itemid")
				
					rsget.moveNext		   
					i=i+1
				loop
			end if
			rsget.close
	end sub	
'##################################################################
	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 11
		FResultCount = 0
		FScrollCount = 11
		FTotalCount =0
	end sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1								'//시작 페이지가 1보다 크면 생성
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1	'//전체 페이지가 시작페이지+전체페이지링크수-1의 수보다 크면 생성
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1	'//시작 페이지는 현재페이지에서 1을 빼고 전체페이지링크수로 나눈후 전체페이지링크수를 곱한후 +1을 하면 생김. 
	end Function
end class

CLASS CAutoCategory
	public FDiscountRate
	public FCategoryList()
	public FCategorySubList()
	public FCategoryPrdList()
	public FCategoryBrand()
	public FItemList()
	public FCategoryPrd
	public FADD()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FResultBCount
	public FScrollCount
	public RoundUP

	public FRectCD1
	public FRectCD2
	public FRectCD3
	
	public FRectBestType

	public FRectCH
	public FRectOrder
	public FRectMakerID
	public FRectStyleGubun
	public FRectItemStyle
	public FRectSort
	public FNotinlist
	Public FRectitemarr
	Public Fdesignerid

	Public FRectOnlySellY
	
	public FRectPercentLow
    public FRectPercentHigh
    
	Private Sub Class_Initialize()
		redim preserve FCategoryList(0)
		redim preserve FCategorySubList(0)
		redim preserve FCategoryPrdList(0)
		redim preserve FCategoryBrand(0)
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 1
		FResultCount = 0
		FResultBCount = 0
		FScrollCount = 10
		FTotalCount =0
			
	
	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	'// 추가 이미지 불러오기 
	Public Sub getAddImage(byval itemid)
			dim strSQL,ArrRows,i
			
			strSQL = "exec [db_item].[dbo].ten_item_addimage_view '" + CStr(itemid) + "'"  + vbcrlf
			
			'rsget.CursorLocation = adUseClient
			'rsget.CursorType=adOpenForwardOnly
			'rsget.Locktype=adLockReadOnly
			rsget.Open strSQL, dbget, 1
			
			If Not rsget.EOF Then 
				ArrRows 	= rsget.GetRows
			End if
			rsget.close
			
			if isArray(ArrRows) then
				
			FResultCount = Ubound(ArrRows,2) + 1
			
			redim  FADD(FResultCount)
			
				For i=0 to FResultCount-1
					Set FADD(i) = new CCategoryPrdItem
					FADD(i).FAddimageGubun	= ArrRows(0,i)
					FADD(i).FAddimageSmall	= "http://webimage.10x10.co.kr/image/add" + Cstr(FADD(i).FAddimageGubun) + "icon/" + GetImageSubFolderByItemid(itemid) + "/C" + ArrRows(1,i)
					FADD(i).FAddimage 			= "http://webimage.10x10.co.kr/image/add" + Cstr(FADD(i).FAddimageGubun) + "/" + GetImageSubFolderByItemid(itemid) + "/" + ArrRows(1,i)
				next 
			end if
	
	
	End Sub
end class	
	
'// 상품 이미지 경로를 계산하여 반환 //
function GetImageSubFolderByItemid(byval iitemid)
    if (iitemid <> "") then
	    GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	else
	    GetImageSubFolderByItemid = ""
	end if
end function
%>	
	
	