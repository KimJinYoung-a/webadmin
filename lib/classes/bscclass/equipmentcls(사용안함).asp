<%
'#######################################################
' Description : 장비 자산 클래스
'	2008년 01월 16일 한용민 생성
'#######################################################
	
class CEquipmentItem
	public Fidx					'// 글번호
	public Fequip_code			'// 장비코드 
	public Fequip_gubun			
	public Fequip_gubun_name	'// 장비구분
	public Fequip_name			'// 제품명
	public Fmodel_name			
	public Fmanufacture_company	
	public Fbuy_company_code	
	public Fbuy_company_name	
	public Fbuy_date			'// 구입한 날짜
	public Fbuy_cost			'// 구입한 비용
	public Fbuy_vat				
	public Fbuy_sum				'// 제품 한개를 구입한 총 금액
	public Fequip_no			
	public Fdurability_month	'// 36개월 
	public Fetc_str				
	public Fdetail_quality1		
	public Fdetail_quality2		
	public Fdetail_qualityetc	
	public Fusinguserid			'// 사용자 이름
	public Fdetail_ip			'// 개인 별 사용하는 ip
	public Fpart_code			'// 사용코드
	public Fpart_code_name		'// 사용구분
	public Fregdate				'// 구매일
	public Flastupdate			
	public Freguserid			
	public Fmodiuserid			
	public fwonga_cost			'구매원가
	public FusinguserName		
	public fstatediv
	public fdel_id
	public fdel_date
	
	public function getDiffDate()'// 구입일 부터 현재까지 경과 개월수
	    If IsDate(Fbuy_date) then
    		if datediff("m",Fbuy_date,Now()) > 0 then
    			getDiffDate = datediff("m", Fbuy_date, Now())
    			'	datediff =		("m", 구입날짜(이전날짜),현재날짜(이후날짜))
    		end if
    	ELSE
    		getDiffDate = 0
    	end if
	end function
	
	'//자산 가치 가격 
	public function getCurrentValue()
		getCurrentValue = 0
		if IsNULL(Fbuy_date)or (Fbuy_date="") then exit function
		getCurrentValue = fwonga_cost  - ((fwonga_cost * getDiffDate)/Fdurability_month)	'정액법 ' 차후 정률법으로 바꾸어야함 1년당 -0.451% 감소
		'현재 페이지의 자산 가치 합계 = 구입가격 - (구입가격 *구입일부터 현재까지의 날짜수/개월수)
	end function
	
	'// 자산가치의 총 합
	public function getAllCurrentValue()	
		dim SQL
																							
		SQL = " select sum(buy_sum-(buy_sum * Datediff(m,buy_date,getdate()))/durability_month) as aaa"		
		SQL = SQL + " from [db_partner].[dbo].tbl_equipment_list"											
		
		'response.write SQL &"<Br>"
		rsget.open SQL, dbget,1																				

		getAllCurrentValue = rsget("aaa")																	
		rsget.close																							
	end function
		
	public function getEquipCode()
		getEquipCode = Fequip_code
	end function

	public function getTotalPrice()
	end function
	
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()

	End Sub
end Class


class CEquipment
	public FOneItem
	public FItemList()
	public FPageSize				
	public FTotalPage				
    public FPageCount					
	public FTotalCount					
	public FTotalSum					'//구입가격의 총 합			
	public FTotalSum2
	public FCurSum						
	public FResultCount					
    public FScrollCount					'//페이지링크 수
	public FCurrPage
	public FRectBuyDateDtStart			'//검색 날짜 시작값
	public FRectBuyDateDtEnd			'//검색 날짜 끝값
	public FRectBuydate					'//제품 구입날짜
	public FRectIdx						'제품 장비코드 
	public FRectJangbi					'장비구분 변수
	public FRectSayoug					'사용구분 변수
	public FRectUser    				
	public FRectIp						'사용중인 ip검색 변수
	public Fequip_code
	public Frectequip_name
	public frectmanufacture_company
	
	public Sub getOneEquipment()
		dim sqlStr, i

		sqlStr = "select top 1 l.*, u.username as usingusername"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_equipment_list l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
		sqlStr = sqlStr + " on l.usinguserid=u.userid"
		sqlStr = sqlStr + " 	and u.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"		
		sqlStr = sqlStr + " where l.idx=" + CStr(FRectIdx)

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		
		FResultCount =  rsget.RecordCount

		set FOneItem = new CEquipmentItem

		i=0
		if  not rsget.EOF  then
			FOneItem.Fidx                 = rsget("idx")
			FOneItem.Fequip_code          = rsget("equip_code")
			FOneItem.Fequip_gubun         = rsget("equip_gubun")
			FOneItem.Fequip_name          = db2html(rsget("equip_name"))
			FOneItem.Fmodel_name          = db2html(rsget("model_name"))
			FOneItem.Fmanufacture_company = db2html(rsget("manufacture_company"))
			FOneItem.Fbuy_company_code    = rsget("buy_company_code")
			FOneItem.Fbuy_company_name    = db2html(rsget("buy_company_name"))
			FOneItem.Fbuy_date            = rsget("buy_date")
			FOneItem.Fbuy_cost            = rsget("buy_cost")
			FOneItem.Fbuy_vat             = rsget("buy_vat")
			FOneItem.Fbuy_sum             = rsget("buy_sum")
			FOneItem.Fequip_no            = rsget("equip_no")
			FOneItem.Fdurability_month    = rsget("durability_month")	'//36개월
			FOneItem.Fdetail_ip			  = rsget("detail_ip")			'//개인별 사용하는 IP
			FOneItem.Fetc_str			  = db2html(rsget("etc_str"))
			FOneItem.Fdetail_quality1	  = db2html(rsget("detail_quality1"))
			FOneItem.Fdetail_quality2	  = db2html(rsget("detail_quality2"))
			FOneItem.Fdetail_qualityetc      = db2html(rsget("detail_qualityetc"))
			FOneItem.Fdetail_ip			  = db2html(rsget("detail_ip"))
			FOneItem.Fusinguserid         = rsget("usinguserid")
			FOneItem.Fpart_code           = rsget("part_code")
			FOneItem.Fregdate             = rsget("regdate")
			FOneItem.Flastupdate          = rsget("lastupdate")
			FOneItem.Freguserid           = rsget("reguserid")
			FOneItem.Fmodiuserid          = rsget("modiuserid")
			FOneItem.FusinguserName		= db2html(rsget("usingusername"))

		end if
		rsget.Close

	end Sub
	
	'//admin/newreport/equipment_list.asp		'//admin/newreport/equipment_excel.asp
	public Sub getEquipmentList()
		dim sqlStr, i, addSQL, sqlSum, ipquery

		addSQL = " where 1=1 "
		
		if Not(FRectJangbi="" or isNull(FRectJangbi)) then						
			addSQL = addSQL + " and l.equip_gubun = '" & FRectJangbi & "'"		
		end if 																	

		if Not(FRectSayoug="" or isNull(FRectSayoug)) then
			addSQL = addSQL + " and l.part_code = '" & FRectSayoug & "'"
		end if
		
		if Not(FRectUser="" or isNull(FRectUser)) then
			addSQL = addSQL + " and l.usinguserid = '" & FRectUser & "'"
		end if
		
		'// Ip 검색 쿼리. 검색 ip에 체크가 없을때 모든 데이터가 나오도록
		if Not(FRectIp="" or isNull(FRectIp)) then											
			addSQL = addSQL + " and l.detail_ip <>''"									
		end if																			
		
		if not(Fequip_code="" or isnull(Fequip_code)) then
			addSQL = addSQL + " and l.equip_code like '%" & Fequip_code & "%'"
		end if

		if Frectequip_name <> "" then
			addSQL = addSQL + " and equip_name like '%" & Frectequip_name & "%'"
		end if			
			
		'// 날짜 검색 쿼리. 검색 날짜 없을때 모든 데이터가 나오도록
		if FRectBuyDateDtStart <>"" then
			addSQL = addSQL + " and buy_date>='" + FRectBuyDateDtStart + "'"
		end if
		
		if FRectBuyDateDtEnd <>"" then
			addSQL = addSQL + " and buy_date<'" + FRectBuyDateDtEnd + "'"	
		end if

		if frectmanufacture_company <>"" then
			addSQL = addSQL + " and manufacture_company like '%" & frectmanufacture_company & "%'"
		end if			
	
		'// 레코드 들의 수를 페이징 하기위해서 쿼리 
		sqlStr = " select count(*) as cnt, sum(buy_sum) as totalprice"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_equipment_list l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
		sqlStr = sqlStr + " 	on l.usinguserid=u.userid"
		sqlStr = sqlStr + " 	and u.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g1"
		sqlStr = sqlStr + " 	on g1.gubuntype='10' and l.equip_gubun=g1.gubuncd"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g2"
		sqlStr = sqlStr + " 	on g2.gubuntype='20' and l.part_code=g2.gubuncd " & addSQL
		'sqlStr = sqlStr + "  left join [db_partner].[dbo].tbl_equipment_list"
		'sqlStr = sqlStr + " 	on l.detail_ip=''" & addSQL
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1										
			FTotalSum = rsget("totalprice")
			FTotalCount = rsget("cnt")
		rsget.Close														
		
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " l.idx,l.equip_code,l.equip_gubun,l.equip_name,l.model_name,l.manufacture_company"
		sqlStr = sqlStr + " ,l.buy_company_code,l.buy_company_name,l.buy_date,l.buy_cost,l.buy_vat"
		sqlStr = sqlStr + " ,l.buy_sum,l.equip_no,l.durability_month,l.detail_quality1,l.detail_quality2"
		sqlStr = sqlStr + " ,l.detail_qualityetc,l.detail_ip,l.etc_str, isnull(l.usinguserid,'') as usinguserid"
		sqlStr = sqlStr + " ,isnull(l.part_code,'') as part_code,l.regdate,l.lastupdate,l.reguserid,l.modiuserid"
		sqlStr = sqlStr + " , u.username as usingusername"
		sqlStr = sqlStr + " , g1.gubunname as equip_gubun_name"
		sqlStr = sqlStr + " , g2.gubunname as part_code_name"
		sqlStr = sqlStr + " ,u.statediv"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_equipment_list l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
		sqlStr = sqlStr + " 	on l.usinguserid=u.userid"
		sqlStr = sqlStr + " 	and u.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g1"
		sqlStr = sqlStr + " 	on g1.gubuntype='10' and l.equip_gubun=g1.gubuncd"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g2"
		sqlStr = sqlStr + " 	on g2.gubuntype='20' and l.part_code=g2.gubuncd " & addSQL
		'sqlStr = sqlStr + "  left join [db_partner].[dbo].tbl_equipment_list"
		'sqlStr = sqlStr + " 	on l.detail_ip=''" & addSQL
		
		if FRectIp<>"" then
			sqlStr = sqlStr + " order by isnull(u.statediv,'Y') asc , l.detail_ip asc"
		else
			sqlStr = sqlStr + " order by isnull(u.statediv,'Y') asc , l.idx desc"		
		end if	

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentItem
				
				FItemList(i).fstatediv                 = rsget("statediv")
				FItemList(i).Fidx                 = rsget("idx")
				FItemList(i).Fequip_code          = rsget("equip_code")
				FItemList(i).Fequip_gubun         = rsget("equip_gubun")
				FItemList(i).Fequip_gubun_name	  = rsget("equip_gubun_name")
				FItemList(i).Fequip_name          = db2html(rsget("equip_name"))
				FItemList(i).Fmodel_name          = db2html(rsget("model_name"))
				FItemList(i).Fmanufacture_company = db2html(rsget("manufacture_company"))
				FItemList(i).Fbuy_company_code    = rsget("buy_company_code")
				FItemList(i).Fbuy_company_name    = db2html(rsget("buy_company_name"))
				FItemList(i).Fbuy_date            = rsget("buy_date")
				FItemList(i).Fbuy_cost            = rsget("buy_cost")
				FItemList(i).fwonga_cost          = rsget("buy_sum") / 1.1				
				FItemList(i).Fbuy_vat             = rsget("buy_vat")
				FItemList(i).Fbuy_sum             = rsget("buy_sum")
				FItemList(i).Fequip_no            = rsget("equip_no")
				FItemList(i).Fdurability_month    = rsget("durability_month")
				FItemList(i).Fdetail_ip			  = rsget("detail_ip")				
				FItemList(i).Fetc_str			  = db2html(rsget("etc_str"))
				FItemList(i).Fdetail_quality1		= db2html(rsget("detail_quality1"))
				FItemList(i).Fdetail_quality2		= db2html(rsget("detail_quality2"))
				FItemList(i).Fdetail_qualityetc     = db2html(rsget("detail_qualityetc"))
				FItemList(i).Fdetail_ip			  	= db2html(rsget("detail_ip"))
				FItemList(i).Fusinguserid         = rsget("usinguserid")
				FItemList(i).Fpart_code           = rsget("part_code")
				FItemList(i).Fpart_code_name      = db2html(rsget("part_code_name"))
				FItemList(i).Fregdate             = rsget("regdate")
				FItemList(i).Flastupdate          = rsget("lastupdate")
				FItemList(i).Freguserid           = rsget("reguserid")
				FItemList(i).Fmodiuserid          = rsget("modiuserid")
				FItemList(i).FusinguserName		= db2html(rsget("usingusername"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub
	
	'//admin/newreport/equipment_loglist.asp
	public Sub getEquipmentlogList()
		dim sqlStr, i, addSQL, sqlSum, ipquery

		if Not(FRectJangbi="" or isNull(FRectJangbi)) then						
			addSQL = addSQL + " and l.equip_gubun = '" & FRectJangbi & "'"		
		end if 																	

		if Not(FRectSayoug="" or isNull(FRectSayoug)) then
			addSQL = addSQL + " and l.part_code = '" & FRectSayoug & "'"
		end if
		
		if Not(FRectUser="" or isNull(FRectUser)) then
			addSQL = addSQL + " and l.usinguserid = '" & FRectUser & "'"
		end if
		
		'// Ip 검색 쿼리. 검색 ip에 체크가 없을때 모든 데이터가 나오도록
		if Not(FRectIp="" or isNull(FRectIp)) then											
			addSQL = addSQL + " and l.detail_ip <>''"									
		end if																			
		
		if not(Fequip_code="" or isnull(Fequip_code)) then
			addSQL = addSQL + " and l.equip_code like '%" & Fequip_code & "%'"
		end if

		if Frectequip_name <> "" then
			addSQL = addSQL + " and equip_name like '%" & Frectequip_name & "%'"
		end if			
			
		'// 날짜 검색 쿼리. 검색 날짜 없을때 모든 데이터가 나오도록
		if FRectBuyDateDtStart <>"" then
			addSQL = addSQL + " and buy_date>='" + FRectBuyDateDtStart + "'"
		end if
		
		if FRectBuyDateDtEnd <>"" then
			addSQL = addSQL + " and buy_date<'" + FRectBuyDateDtEnd + "'"	
		end if
		
		'// 레코드 들의 수를 페이징 하기위해서 쿼리 
		sqlStr = " select count(*) as cnt, sum(buy_sum) as totalprice"
		sqlStr = sqlStr + " from [db_partner].dbo.tbl_equipment_log l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
		sqlStr = sqlStr + " 	on l.usinguserid=u.userid"
		sqlStr = sqlStr + " 	and u.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g1"
		sqlStr = sqlStr + " 	on g1.gubuntype='10' and l.equip_gubun=g1.gubuncd"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g2"
		sqlStr = sqlStr + " 	on g2.gubuntype='20' and l.part_code=g2.gubuncd " & addSQL
		'sqlStr = sqlStr + "  left join [db_partner].[dbo].tbl_equipment_list"
		'sqlStr = sqlStr + " 	on l.detail_ip=''" & addSQL
		sqlStr = sqlStr + " where 1=1 " & addSQL
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1										
			FTotalSum = rsget("totalprice")
			FTotalCount = rsget("cnt")
		rsget.Close														
		
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " l.idx,l.equip_code,l.equip_gubun,l.equip_name,l.model_name,l.manufacture_company"
		sqlStr = sqlStr + " ,l.buy_company_code,l.buy_company_name,l.buy_date,l.buy_cost,l.buy_vat"
		sqlStr = sqlStr + " ,l.buy_sum,l.equip_no,l.durability_month,l.detail_quality1,l.detail_quality2"
		sqlStr = sqlStr + " ,l.detail_qualityetc,l.detail_ip,l.etc_str, isnull(l.usinguserid,'') as usinguserid"
		sqlStr = sqlStr + " ,isnull(l.part_code,'') as part_code,l.regdate,l.lastupdate,l.reguserid,l.modiuserid"
		sqlStr = sqlStr + " ,l.del_id, l.del_date"
		sqlStr = sqlStr + " , u.username as usingusername"
		sqlStr = sqlStr + " , g1.gubunname as equip_gubun_name"
		sqlStr = sqlStr + " , g2.gubunname as part_code_name"
		sqlStr = sqlStr + " ,u.statediv"
		sqlStr = sqlStr + " from [db_partner].dbo.tbl_equipment_log l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
		sqlStr = sqlStr + " 	on l.usinguserid=u.userid"
		sqlStr = sqlStr + " 	and u.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g1"
		sqlStr = sqlStr + " 	on g1.gubuntype='10' and l.equip_gubun=g1.gubuncd"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun g2"
		sqlStr = sqlStr + " 	on g2.gubuntype='20' and l.part_code=g2.gubuncd " & addSQL
		'sqlStr = sqlStr + "  left join [db_partner].[dbo].tbl_equipment_list"
		'sqlStr = sqlStr + " 	on l.detail_ip=''" & addSQL
		sqlStr = sqlStr + " where 1=1 " & addSQL
		sqlStr = sqlStr + " order by l.idx desc"

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentItem
				
				FItemList(i).fdel_id                 = rsget("del_id")
				FItemList(i).fdel_date                 = rsget("del_date")
				FItemList(i).fstatediv                 = rsget("statediv")
				FItemList(i).Fidx                 = rsget("idx")
				FItemList(i).Fequip_code          = rsget("equip_code")
				FItemList(i).Fequip_gubun         = rsget("equip_gubun")
				FItemList(i).Fequip_gubun_name	  = rsget("equip_gubun_name")
				FItemList(i).Fequip_name          = db2html(rsget("equip_name"))
				FItemList(i).Fmodel_name          = db2html(rsget("model_name"))
				FItemList(i).Fmanufacture_company = db2html(rsget("manufacture_company"))
				FItemList(i).Fbuy_company_code    = rsget("buy_company_code")
				FItemList(i).Fbuy_company_name    = db2html(rsget("buy_company_name"))
				FItemList(i).Fbuy_date            = rsget("buy_date")
				FItemList(i).Fbuy_cost            = rsget("buy_cost")
				FItemList(i).fwonga_cost          = rsget("buy_sum") / 1.1				
				FItemList(i).Fbuy_vat             = rsget("buy_vat")
				FItemList(i).Fbuy_sum             = rsget("buy_sum")
				FItemList(i).Fequip_no            = rsget("equip_no")
				FItemList(i).Fdurability_month    = rsget("durability_month")
				FItemList(i).Fdetail_ip			  = rsget("detail_ip")				
				FItemList(i).Fetc_str			  = db2html(rsget("etc_str"))
				FItemList(i).Fdetail_quality1		= db2html(rsget("detail_quality1"))
				FItemList(i).Fdetail_quality2		= db2html(rsget("detail_quality2"))
				FItemList(i).Fdetail_qualityetc     = db2html(rsget("detail_qualityetc"))
				FItemList(i).Fdetail_ip			  	= db2html(rsget("detail_ip"))
				FItemList(i).Fusinguserid         = rsget("usinguserid")
				FItemList(i).Fpart_code           = rsget("part_code")
				FItemList(i).Fpart_code_name      = db2html(rsget("part_code_name"))
				FItemList(i).Fregdate             = rsget("regdate")
				FItemList(i).Flastupdate          = rsget("lastupdate")
				FItemList(i).Freguserid           = rsget("reguserid")
				FItemList(i).Fmodiuserid          = rsget("modiuserid")
				FItemList(i).FusinguserName		= db2html(rsget("usingusername"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub
	
	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub
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

Sub DrawipGubun(selectBoxName)
   dim tmp_str,query1
   %>
   <select onChange=javascript:checkip(this); name="checkipform">
   <option value=''>사용가능IP(선택)</option><%
   query1 = " select company_ip from [db_partner].[dbo].tbl_equipment_ip where company_name is null or company_name=''"
   query1 = query1 + " order by company_ip Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       'rsget.Movefirst

       do until rsget.EOF
           'if Lcase(selectedId) = Lcase(rsget("gubuncd")) then
               'tmp_str = " selected"
          ' end if
			response.write "<option value='"&rsget("company_ip")&"'>" & rsget("company_ip") & "</option>" 
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub DrawipGubun2(selectBoxName)
   dim tmp_str,query1
   %>
   <select>
   <option value=''>회사내 사용중인 IP</option><%
   query1 = " select company_ip , company_name from [db_partner].[dbo].tbl_equipment_ip where company_name<>''"
   query1 = query1 + " order by company_ip Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       'rsget.Movefirst

       do until rsget.EOF
           'if Lcase(selectedId) = Lcase(rsget("gubuncd")) then
           '    tmp_str = " selected"
           'end if
			response.write "<option>" & rsget("company_name") & "(" & rsget("company_ip") & ") </option>" 
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

'셀렉트 옵션 생성 함수(장비구분, 사용구분)
Sub DrawEquipMentGubun(gubuntype,selectBoxName,selectedId,chplg)		
   dim tmp_str,query1, qyery2									

	response.write "<select name='" & selectBoxName & "' "&chplg&">"		
	response.write "<option value=''"							
		if selectedId="" then									
			response.write " selected"
		end if
	response.write ">선택</option>"								

	 '옵션 내용 DB에서 가져오기
   query1 = " select gubuncd,gubunname from [db_partner].[dbo].tbl_equipment_gubun where gubuntype='" + gubuntype + "'"
   query1 = query1 + " order by gubuncd"
   rsget.Open query1,dbget,1									

   if  not rsget.EOF  then										

       '도돌이 시작
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("gubuncd")) then 		
               tmp_str = " selected"					
           end if
           response.write("<option value='"&rsget("gubuncd")&"' "&tmp_str&">" + db2html(rsget("gubunname")) + "</option>")
           tmp_str = ""					
           rsget.MoveNext
       loop
   end if
   rsget.close

   '셀렉트 끝
   response.write("</select>")
End Sub

'//셀렉트 옵션 생성 함수(사용자 검색)
Sub DrawUserGubun(selectboxname, usinguserid)		
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "'>"		
	response.write "<option value=''"						
		if usinguserid ="" then								
			response.write "selected"
		end if
	response.write ">선택</option>"							

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select usinguserid from [db_partner].[dbo].tbl_equipment_list"
	userquery = userquery + " where usinguserid<>'' "
	userquery = userquery + " group by usinguserid " 'group by
	userquery = userquery + " order by usinguserid desc"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(usinguserid) = Lcase(rsget("usinguserid")) then 
				tem_str = " selected"							
			end if
			response.write "<option value='" & rsget("usinguserid") & "' " & tem_str & ">" & db2html(rsget("usinguserid")) & "</option>"
			tem_str = ""			
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub

Sub drawpartneruser(byval selectBoxName, selectedId ,chplg)
   dim tmp_str,sqlStr ,tmp_substr

	sqlStr = "select"
	sqlStr = sqlStr & " pi.part_name, t.empno  , t.username ,t.userid ,t.statediv"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and t.isUsing = 1"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi"
	sqlStr = sqlStr & " 	on t.part_sn = pi.part_sn"
	sqlStr = sqlStr & " 	and pi.part_isdel = 'N'"
	sqlStr = sqlStr & " order by t.statediv desc ,t.part_sn desc, t.posit_sn asc ,t.username asc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	%>
	<select class='select' name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<%
   
		if not rsget.EOF then
			rsget.Movefirst
		
			do until rsget.EOF
		
				tmp_substr = ""
				
				if selectedId <> "" then
					if selectedId = rsget("userid") then
						tmp_str = " selected"
					end if
				end if
				
				tmp_substr = tmp_substr + db2html(rsget("part_name")) + "-"
				tmp_substr = tmp_substr + db2html(rsget("username"))
				
				if rsget("userid") <> "" then tmp_substr = tmp_substr + " (" + rsget("userid") + ")"
				
				if rsget("statediv") <> "Y" then tmp_substr = tmp_substr + " (퇴사)"
					
				response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + tmp_substr + "</option>")
				tmp_str = ""
				rsget.MoveNext
			loop
		end if
	rsget.close
	response.write("</select>")
end Sub
%>