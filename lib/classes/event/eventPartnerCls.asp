<% 

public Function fnSetThemeNm(ByVal TCode)
dim themeNm
SELECT CASE TCode
	CASE "2"
		themeNm = "이미지 테마"
	CASE "3"
		themeNm = "상품 테마"
	CASE ELSE
		themeNm = "텍스트 테마"
	END SELECT
	
	fnSetThemeNm = themeNm	
End Function

public Function fnSetStatusNm(ByVal SCode)
dim statusNm
SELECT CASE SCode
	CASE "5"
		statusNm = "<span class='tag bgYw1'>승인요청</span>"
	CASE "7"
		statusNm = "<span class='tag bgGn1'>승인</span>"
	CASE "6"
		statusNm = "<span class='tag bgBl1'>오픈</span>"
	CASE "3"
		statusNm = "<span class='tag bgRd1'>반려</span>"
	CASE "9"
		statusNm = "<span class='tag bgGy1'>종료</span>"				
	CASE ELSE
		statusNm = "<span class='tag bgYw1'>등록</span>"
END SELECT

fnSetStatusNm  = statusNm
End Function

public Function fnSetTextForm(ByVal sValue)
if sValue ="" or isNull(sValue) then exit function
fnSetTextForm = replace(sValue,chr(13),"<br>")
end Function


	public Function fnEventColorCode(ByVal Fthemecolor)
		 
			If Fthemecolor="1" Then
				fnEventColorCode = "#ed6c6c"
			ElseIf Fthemecolor="2" Then
				fnEventColorCode = "#f385af"
			ElseIf Fthemecolor="3" Then
				fnEventColorCode = "#f3a056"
			ElseIf Fthemecolor="4" Then
				fnEventColorCode = "#e7b93c"
			ElseIf Fthemecolor="5" Then
				fnEventColorCode = "#8eba4a"
			ElseIf Fthemecolor="6" Then
				fnEventColorCode = "#43a251"
			ElseIf Fthemecolor="7" Then
				fnEventColorCode = "#50bdd1"
			ElseIf Fthemecolor="8" Then
				fnEventColorCode = "#5aa5ea"
			ElseIf Fthemecolor="9" Then
				fnEventColorCode = "#2672bf"
			ElseIf Fthemecolor="10" Then
				fnEventColorCode = "#2c5a85"
			ElseIf Fthemecolor="11" Then
				fnEventColorCode = "#848484"
			Else
				fnEventColorCode = "#848484"
			End If
	 
	End Function

	public Function fnEventBarColorCode(ByVal Fthemecolor)
		 
			If Fthemecolor="1" Then
				fnEventBarColorCode = "#cb4848"
			ElseIf Fthemecolor="2" Then
				fnEventBarColorCode = "#d55787"
			ElseIf Fthemecolor="3" Then
				fnEventBarColorCode = "#e37f35"
			ElseIf Fthemecolor="4" Then
				fnEventBarColorCode = "#ce8d00"
			ElseIf Fthemecolor="5" Then
				fnEventBarColorCode = "#699426"
			ElseIf Fthemecolor="6" Then
				fnEventBarColorCode = "#358240"
			ElseIf Fthemecolor="7" Then
				fnEventBarColorCode = "#2899ae"
			ElseIf Fthemecolor="8" Then
				fnEventBarColorCode = "#2f7cc3"
			ElseIf Fthemecolor="9" Then
				fnEventBarColorCode = "#145290"
			ElseIf Fthemecolor="10" Then
				fnEventBarColorCode = "#1c3e5d"
			ElseIf Fthemecolor="11" Then
				fnEventBarColorCode = "#656565"
			Else
				fnEventBarColorCode = "#656565"
			End If
		 
		 
	End Function
	
'================================================================================================
Class CEvent

 public FevtCode
 public Fmakerid 		
 public Fevtkind      
 public Fevtmanager   
 public Fevtname      
 public Fevtstartdate 
 public Fevtenddate   
 public Fevtstate     
 public Fevtregdate   
 public Fevtusing     
 public Fevtlastupdate
 public Fadminid      
            
 public Fevtcategory  
 public FevtcateMid   
 public FevtCateNm
 public FevtCateMNm
 public Fissale       
 public Fisgift       
 public Fiscoupon     
 public Fbrand        
 public Fevttag       

public FevtGCode

public FsalePer
public FsaleCPer
public FbrandNm 
public FTitlePC
public FTitleMO
public Fetcitemimg 
public Fevt_mo_listbanner 
public FsubcopyK
public Fevtsubname
public Fmdtheme
public Fthemecolor
public Fthemecolormo
public Ftextbgcolor
public Fgiftisusing
public Fgifttext1
public Fgiftimg1
public Fgifttext2
public Fgiftimg2
public Fgifttext3
public Fgiftimg3
public FSdiv
public Fevtdispcate

 public FRectmakerid
 public FRectSType
 public FRectSDate
 public FRectEDate
 public FRectUsing
 public FRectState
 public FRectDisp1
 public FRectDisp2
 public FRectDispcate
 public FRectNm 
 public FRectECode
 public FRectevtstate
 public FRealECode
 public FRealEState
 
 
public FRectItemid
public FRectItemName
public FPageSize
public FCurrPage
public FRectSailYn
public FRectSort
public FRectSellYN 

public FitemTotCnt
 public FTotCnt
 public FPSize
 public FCPage
 public FSPageNo
 public FEPageNo
 public FTotalPage
 
 
 ''//통합 상태값 확인
 public Function fnGetTotState
 
  dim strSql
  strSql ="select   evt_state , realevt_code,  case when realevt_code is not null then (select evt_state from db_event.dbo.tbl_event where evt_code = pe.realevt_code ) else 0 end as realevtstate "
  strSql= strSql &" from db_Event.dbo.tbl_partner_event as pe where evt_code ="&FRectECode
  rsget.Open strSql,dbget,1
  if not rsget.eof  then
  	 FRectevtstate = rsget("evt_state")
  	 FRealECode= rsget("realevt_code")
  	 FRealEState = rsget("realevtstate")
 end if
rsget.close
end Function

 ''// 리스트
 public Function fnGetEventList
  dim strSql
  dim strSearch
  strSearch = ""
  if  FRectSDate <> "" and FRectEDate <> "" then
  		if FRectSType ="1"   then  '//시작일
  			strSearch = strSearch & " and evt_startdate >= '"&FRectSDate&"' and evt_startdate <= '"&FRectEdate&"'"
  		elseif FRectSType ="2"   then  '//종료일
  			strSearch = strSearch & " and evt_enddate >= '"&FRectSDate&"' and evt_enddate <= '"&FRectEdate&"'"
  		elseif FRectSType ="3"   then  '//작성일
  			strSearch = strSearch & " and evt_regdate >= '"&FRectSDate&"' and evt_regdate <= '"&FRectEdate&"'"	
  		end if
end if	
  	
  if FRectUsing <> "" and FRectUsing <> "A"	then 
  	strSearch = strSearch & " and evt_using ='"&FRectUsing&"'"
  end if
  
  if FRectState <> "" and 	FRectState <> "A" then  
  	strSearch = strSearch & " and evt_state in ("&FRectState&")"
  end if
   
  if FRectDispcate <> "" then
  	strSearch = strSearch & "  and left(evt_dispcate,"&len(FRectDispcate)&")="&FRectDispcate
end if
 

  if FRectNm <> "" then
  	strSearch = strSearch & "  and evt_name like '%"&FRectNm&"%'"
end if
  
  	strSql = "  select  evt_code, revt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using, evt_dispcate, brand,salePer,saleCPer,mdtheme 	"&vbCrlf
	strSql = strSql & "  into #evtList "&vbCrlf
	strSql = strSql & " from ( "&vbCrlf
	strSql = strSql & " 	select evt_code , 0 as revt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using, evt_dispcate, brand,salePer,saleCPer,mdtheme 	"&vbCrlf
	strSql = strSql & "  	from db_event.dbo.tbl_partner_event "&vbCrlf
	strSql = strSql & " 	where brand ='"&FRectmakerid&"'  and evt_state <=5 "&vbCrlf
	strSql = strSql & " 	union all"&vbCrlf
	strSql = strSql & " 	select  p.evt_code, e.evt_code as revt_code,e.evt_name,e.evt_startdate,e.evt_enddate,case when e.evt_state <6 then 7 else e.evt_state end as evt_state,e.evt_regdate,e.evt_using, d.evt_dispcate, d.brand,d.salePer,d.saleCPer,d.mdtheme "&vbCrlf
	strSql = strSql & " 	from db_event.dbo.tbl_partner_event as p "&vbCrlf
	strSql = strSql & " 	inner join db_event.dbo.tbl_event as e on p.realevt_code = e.evt_code "&vbCrlf
	strSql = strSql & "  	inner join db_event.dbo.tbl_event_display as d on e.evt_code = d.evt_code "&vbCrlf
	strSql = strSql & " 	where  p.brand ='"&FRectmakerid&"' and p.realevt_code is not null and p.evt_state >5 "&vbCrlf
	strSql = strSql & " 	) as t  "&vbCrlf
	strSql = strSql & "	order by evt_regdate desc "&vbCrlf
	rsget.Open strSql,dbget,1
		
  strSql = " SELECT count(evt_code)  FROM #evtList   "
  rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
	rsget.close

if 		FTotCnt > 0 then	
	FSPageNo = (FPSize*(FCPage-1)) + 1
	FEPageNo = FPSize*FCPage	
	
	
  strSql = "SELECT  evt_code, revt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using, evt_dispcate,brand,salePer,saleCPer,mdtheme ,dc1nm, dc2nm "&vbCrlf
  strSql = strSql & " FROM ( "&vbCrlf
  strSql = strSql & "  	SELECT ROW_NUMBER() OVER (ORDER BY evt_regdate desc ) as RowNum "&vbCrlf
  strSql = strSql & "			,evt_code,revt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,evt_dispcate,brand,salePer,saleCPer,mdtheme ,dc1.catename as dc1nm, dc2.catename as dc2nm "&vbCrlf
  strSql = strSql & " 		FROM #evtList as e "&vbCrlf
  strSql = strSql & "			left outer join db_item.dbo.tbl_display_cate as dc1 on    dc1.catecode = left(e.evt_dispcate ,3)"&vbCrlf
  strSql = strSql & " 			 left outer  join db_item.dbo.tbl_display_cate as dc2 on  dc2.catecode =  e.evt_dispcate  "&vbCrlf
  strSql = strSql &"   	where brand = '"&FRectmakerid&"'" & strSearch&vbCrlf
  strSql = strSql & " ) as TB "&vbCrlf
  strSql = strSql & " WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo &vbCrlf
 strSql = strSql & " order by evt_regdate desc "
   rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventList = rsget.getRows()
	end if
	rsget.close
end if	
End Function

'//1단계 등록내용
 public Function fnGetEventST1
 	dim strSql
 	strSql = "select evt_code,evt_kind,evt_manager,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,evt_lastupdate,adminid,evt_dispcate,issale,isgift,iscoupon,brand,evt_tag "
 	strSql = strSql & " from db_Event.dbo.tbl_partner_event where evt_code = "&FevtCode&" and brand = '"&Fmakerid&"'"
 	 
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 	  FevtCode 			= rsget("evt_code")
 	  Fevtkind     	 = rsget("evt_kind")       
 	  Fevtmanager   = rsget("evt_manager")    
 	  Fevtname       = rsget("evt_name")       
 	  Fevtstartdate  = rsget("evt_startdate")  
 	  Fevtenddate   = rsget("evt_enddate")    
 	  Fevtstate        = rsget("evt_state")      
 	  Fevtregdate     = rsget("evt_regdate")    
 	  Fevtusing        = rsget("evt_using")      
 	  Fevtlastupdate = rsget("evt_lastupdate") 
 	  Fadminid           = rsget("adminid")        
 	                                         
 	  Fevtdispcate   = rsget("evt_dispcate")    
 	  Fissale             = rsget("issale")         
 	  Fisgift               = rsget("isgift")         
 	  Fiscoupon         = rsget("iscoupon")       
 	  Fbrand              = rsget("brand")          
 	  Fevttag            = rsget("evt_tag")       
 	end if                                 
	rsget.close                            
 End Function    


'//2단계 그룹리스트
public Function fnGetEventGroup
 	dim strSql 
 	if FRectevtstate = "" THEN FRectevtstate = 0
 	if FRectevtstate > 5 then  	
 		strSql = "select g.evtgroup_code, g.evtgroup_desc, g.evtgroup_sort "
	 	strSql = strSql & ",(select count(itemid) from db_event.dbo.tbl_eventitem where evt_code = g.evt_code and evtgroup_code = g.evtgroup_code and evtitem_isUsing =1) as evtitem_cnt"
	 	strSql = strSql & " from db_event.dbo.tbl_eventitem_group as g "
	 	strSql = strSql & "	inner join db_event.dbo.tbl_event as e on g.evt_code = e.evt_code  "
	 	strSql = strSql & " where g.evt_code="&FevtCode&" and g.evtgroup_using ='Y' and g.evtgroup_pcode >0 order by g.evtgroup_sort " 
	else
		strSql = "select g.evtgroup_code, g.evtgroup_desc, g.evtgroup_sort "
	 	strSql = strSql & ",(select count(itemid) from db_event.dbo.tbl_partner_eventitem where evt_code = g.evt_code and evtgroup_code = g.evtgroup_code and evtitem_isUsing =1) as evtitem_cnt"
	 	strSql = strSql & " from db_event.dbo.tbl_partner_eventitem_group as g "
	 	strSql = strSql & "	inner join db_event.dbo.tbl_partner_event as e on g.evt_code = e.evt_code  and e.brand = '"&Fmakerid&"'"
	 	strSql = strSql & " where g.evt_code="&FevtCode&" and g.evtgroup_using ='Y' and g.evtgroup_pcode >0  order by g.evtgroup_sort " 
	end if 	
 	 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventGroup = rsget.getRows()
	end if
	rsget.close
End function
 
 
'//2단계 그룹리스트
public Function fnGetPVEventGroup
 	dim strSql 
 	 
		strSql = "select g.evtgroup_code, g.evtgroup_desc, g.evtgroup_sort "
	 	strSql = strSql & ",(select count(itemid) from db_event.dbo.tbl_partner_eventitem where evt_code = g.evt_code and evtgroup_code = g.evtgroup_code and evtitem_isUsing =1) as evtitem_cnt"
	 	strSql = strSql & " from db_event.dbo.tbl_partner_eventitem_group as g "
	 	strSql = strSql & "	inner join db_event.dbo.tbl_partner_event as e on g.evt_code = e.evt_code  and e.brand = '"&Fmakerid&"'"
	 	strSql = strSql & " where g.evt_code="&FevtCode&" and g.evtgroup_using ='Y'   order by g.evtgroup_sort " 
	 	
 	 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetPVEventGroup = rsget.getRows()
	end if
	rsget.close
End function

'//2단계 그룹 등록상품 리스트
public Function fnGetEventGroupItem
	dim strSql ,sDBNm
	if FRectevtstate = "" THEN FRectevtstate = 0
 	if FRectevtstate > 5 then   
		sDBNm="db_event.[dbo].[tbl_eventitem]"
	else
		sDBNm="db_event.[dbo].[tbl_partner_eventitem]"
end if
	strSql ="SELECT count(itemid) FROM "&sDBNm&" where evt_code ="&FevtCode&" and evtgroup_code ="&FevtGCode& " and evtitem_isusing = 1 "
		rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
 	rsget.close
 	
 	if FTotCnt >0 then
	strSql = "SELECT  ei.[itemid],i.itemname, ei.[evtitem_sort],ei.[evtitem_imgsize],sailyn, sellcash, buycash, orgprice, orgsuplycash, sailprice, sailsuplycash,   itemcouponyn, itemcoupontype, itemcouponvalue, sellyn "
  	strSql = strSql & " FROM "&sDBNm&" as ei "  	
  	strSql = strSql & " INNER JOIN db_item.dbo.tbl_item as i on ei.itemid = i.itemid "
  	strSql = strSql & " WHERE ei.evt_code ="&FevtCode&" and ei.evtgroup_code ="&FevtGCode&" and evtitem_isusing = 1 "
  	strSql = strSql & "	order by ei.evtitem_sort "
  	 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventGroupItem = rsget.getRows()
	end if
	rsget.close
	end if
End Function

'//2단계 등록내용
public Function fnGetEventST2
	dim strSql
	strSql = " select issale, iscoupon, saleper, salecper from db_Event.dbo.tbl_partner_event where evt_code = "&FevtCode&" and brand = '"&Fmakerid&"'"
	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 	  Fissale             = rsget("issale")          	  
 	  Fiscoupon         = rsget("iscoupon")    
 	  FsalePer 			= rsget("saleper")    
 	  FsaleCPer 			= rsget("salecper")    
	end if
	rsget.close
End function


'//3단계 등록내용
 public Function fnGetEventST3
 	dim strSql
 	strSql = "select socname, etc_itemimg, evt_mo_listbanner ,mdtheme, title_pc, title_mo, evt_subcopyK,evt_subname, issale, iscoupon, saleper, salecper,themecolor,themecolormo,textbgcolor "
 	strSql = strSql & " ,gift_isusing,gift_text1,gift_img1,gift_text2,gift_img2,gift_text3,gift_img3 "
 	strSql = strSql & " from db_Event.dbo.tbl_partner_event as e "
 	strSql = strSql & " inner join  db_user.dbo.tbl_user_c as c on e.brand = c.userid  "
 	strSql= strSql & "  where evt_code = "&FevtCode&" and brand = '"&Fmakerid&"'" 	 
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 	 FbrandNm = rsget("socname")
 	  FTitlePC= rsget("title_pc")
 	  FTitleMO= rsget("title_mo")
 	  Fissale             = rsget("issale")          	  
 	  Fiscoupon         = rsget("iscoupon")    
 	  FsalePer 			= rsget("saleper")    
 	  FsaleCPer 			= rsget("salecper")    
	 Fetcitemimg          = rsget("etc_itemimg") 
	 Fevt_mo_listbanner  = rsget("evt_mo_listbanner")  
	 FsubcopyK           = rsget("evt_subcopyK")
	 Fevtsubname         = rsget("evt_subname")
	 Fmdtheme            = rsget("mdtheme")
	 Fthemecolor         = rsget("themecolor")
	 Fthemecolormo       = rsget("themecolormo")
	 Ftextbgcolor        = rsget("textbgcolor")
	 Fgiftisusing        = rsget("gift_isusing")
	 Fgifttext1          = rsget("gift_text1")
	 Fgiftimg1           = rsget("gift_img1")
	 Fgifttext2          = rsget("gift_text2")
	 Fgiftimg2           = rsget("gift_img2")
	 Fgifttext3          = rsget("gift_text3")
	 Fgiftimg3          = rsget("gift_img3")
	end if                               
	rsget.close                            
 End Function                           
 
 
'//상품배너 선택상품
public Function fnGetProductList
        dim sqlStr, i

        sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr &"	inner join db_event.dbo.tbl_partner_eventitem as e on i.itemid = e.itemid "
        sqlStr = sqlStr & " where i.isusing ='Y' and i.itemid<>0 and e.evt_code = "&FevtCode&" and e.evtitem_isusing=1 "
        if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
							FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if
        if (FRectItemName <> "") then
            '[ 검색이 안된다며 -play auto [ => [[] : 임시상품코드 실상품코드 매핑 위해 상품명조회가 필요
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if
		if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if 
'        if (FRectMWDiv="MW") then
'            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
'        elseif (FRectMWDiv<>"") then
'            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
'        end if

'		if (FRectLimityn="Y0") then
'            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
'        elseif (FRectLimityn<>"") then
'            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
'        end if 
		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if            
		''딜상품 제외
		sqlStr = sqlStr + " and i.itemdiv<>'21'" 
        rsget.Open sqlStr,dbget,1
            FitemTotCnt = rsget("cnt")
        rsget.Close

 
 				FSPageNo = (FPSize*(FCPage-1)) + 1
				FEPageNo = FPSize*FCPage	
	
				sqlStr = " select itemid, makerid, itemname, sellcash, buycash, sellyn, isusing, mwdiv, limityn, limitno,limitsold,regdate ,imgsmall,upchemanagecode, deliverytype "
				sqlStr = sqlStr & ", orgprice, orgsuplycash, sailprice,sailsuplycash,sailyn,itemcouponyn,curritemcouponidx,itemcoupontype,itemcouponvalue,couponbuyprice ,deliverOverseas"
				sqlStr = sqlStr & " from ( "
        sqlStr = sqlStr & "select  ROW_NUMBER() OVER (ORDER BY i.itemid desc ) as RowNum "    
        sqlStr = sqlStr & " ,i.itemid, i.makerid, i.itemname, i.sellcash, i.buycash, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold "
        sqlStr = sqlStr & " ,i.regdate, IsNull(i.smallimage,'') as imgsmall "
        sqlStr = sqlStr & " , isNull(i.upchemanagecode,'') as upchemanagecode, i.deliverytype "       
				sqlStr = sqlStr & " ,i.orgprice, i.orgsuplycash, i.sailprice,i.sailsuplycash,i.sailyn,i.itemcouponyn,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue"
				sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
				sqlStr = sqlStr & ", deliverOverseas "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr &"	inner join db_event.dbo.tbl_partner_eventitem as e on i.itemid = e.itemid "
        sqlStr = sqlStr & " where i.isusing ='Y'  and e.evt_code = "&FevtCode&" and e.evtitem_isusing=1 " 
        sqlStr = sqlStr & " and i.itemid<>0"

       if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemName <> "") then
            sqlStr = sqlStr & " and i.itemname like '%" + html2db(replace(FRectItemName,"[","[[]")) + "%'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
							FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectSellYN="YS") then
            sqlStr = sqlStr & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlStr = sqlStr & " and i.sellyn='" + FRectSellYN + "'"
        end if
 

'        if (FRectMWDiv="MW") then
'            sqlStr = sqlStr + " and (i.mwdiv='M' or i.mwdiv='W')"
'        elseif (FRectMWDiv<>"") then
'            sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
'        end if
'
'		if (FRectLimityn="Y0") then
'            sqlStr = sqlStr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
'        elseif (FRectLimityn<>"") then
'            sqlStr = sqlStr + " and i.limityn='" + FRectLimityn + "'"
'        end if

     
		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if
		
		if FRectSailYn <> "" then '20161103 추가
				sqlStr = sqlStr + " and i.sailyn = '"+FRectSailYn +"'"
		end if
      
        '딜상품 제외
		sqlStr = sqlStr + " and i.itemdiv<>'21'"
		sqlStr = sqlStr + " ) as TB "
		sqlStr = sqlStr + " WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo 
 		sqlStr = sqlStr + " order by itemid desc "  
        rsget.Open sqlStr,dbget,1 

        if not rsget.eof then
        	 fnGetProductList = rsget.getRows()
        end if
        rsget.Close
    end Function


 '//3단계 상품배너 
 public Function fnGetEventItemBanner
 dim strSql,sdbNm
 if FRectevtstate = "" THEN FRectevtstate = 0
 	if FRectevtstate > 5 then  
	sdbNm ="[db_event].[dbo].[tbl_event_itembanner]"
else
	sdbNm ="[db_event].[dbo].[tbl_partner_event_itembanner]"
end if	

 strSql = "select count(b.idx) "
 strSql = strSql & " from "&sdbNm&" as b "
 strSql = strSql & " inner join db_item.dbo.tbl_item as i on b.itemid = i.itemid "
 strSql = strSql & " where b.evt_code = "&FevtCode&"  and b.sdiv='"&FSdiv&"'" 	 
 rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
 	rsget.close
 	
 strSql = "select b.idx, b.itemid, b.viewidx, i.itemname,sailyn, sellcash, buycash, orgprice, orgsuplycash, sailprice, sailsuplycash,   itemcouponyn, itemcoupontype, itemcouponvalue"
 strSql = strSql & " from "&sdbNm&" as b "
 strSql = strSql & " inner join db_item.dbo.tbl_item as i on b.itemid = i.itemid "
 strSql = strSql & " where b.evt_code = "&FevtCode&"  and b.sdiv='"&FSdiv&"'" 	 
 strSql = strSql & " order by b.viewidx"
 
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventItemBanner = rsget.getrows()
	end if
	rsget.close
End Function 

'//3단계 상품배너 이미지
public Function fnGetEventItemImg
dim sqlStr,sdbNm
if FRectevtstate = "" THEN FRectevtstate = 0
 if FRectevtstate > 5 then  
	sdbNm ="[db_event].[dbo].[tbl_event_itembanner]"
else
	sdbNm ="[db_event].[dbo].[tbl_partner_event_itembanner]"
end if	

	sqlStr = "select  top 3  i.basicimage, i.itemid ,i.listimage "
	sqlStr = sqlStr & " from   "&sdbNm&" as e "
	sqlStr = sqlStr & "	join [db_item].[dbo].tbl_item i on e.itemid = i.itemid "
	sqlStr= sqlStr &" where e.evt_code= '" & FevtCode & "'   and e.sdiv='"&FSdiv&"' "
	sqlStr = sqlStr & " order by e.viewidx asc"
	 
	rsget.Open sqlStr,dbget,1
 	if not rsget.eof then
 		fnGetEventItemImg = rsget.getrows()
	end if
	rsget.close
End Function

'// 이벤트 모바일 추가 배너
public Function fnGetMoSlideImgCnt()
	Dim strSql
	If FevtCode = "" THEN Exit Function
	strSql ="SELECT count(idx) as cnt FROM [db_event].[dbo].[tbl_partner_event_slide_addimage] where  evt_code="&FevtCode&" and device='M' and isusing='Y'"
	rsget.Open strSql,dbget,1
	IF Not (rsget.EOF OR rsget.BOF) THEN
		fnGetMoSlideImgCnt = rsget("cnt")
	END IF
	rsget.close
End Function

'// 이벤트 모바일 추가 배너
public Function fnGetMoItemSlideImgCnt()
	Dim strSql
	If FevtCode = "" THEN Exit Function
	strSql ="SELECT count(idx) as cnt FROM [db_event].[dbo].[tbl_partner_event_itembanner] where  evt_code="&FevtCode&" and sdiv='m'"
	rsget.Open strSql,dbget,1
	IF Not (rsget.EOF OR rsget.BOF) THEN
		fnGetMoItemSlideImgCnt = rsget("cnt")
	END IF
	rsget.close
End function 
  
 '//3단계 이미지
 public Function fnGetEventSlideImg
dim sqlStr, sdbNm
if FRectevtstate = "" THEN FRectevtstate = 0
 	if FRectevtstate > 5 then  
	sdbNm ="[db_event].[dbo].[tbl_event_slide_addimage]"
else
	sdbNm ="[db_event].[dbo].[tbl_partner_event_slide_addimage]"
end if	
	sqlStr = "select  top 3 slideimg, idx, device  "
	sqlStr = sqlStr & " from   "&sdbNm
	sqlStr= sqlStr &" where evt_code= '" & FevtCode & "'   and device='"&FSdiv&"' and isusing ='Y' "
	sqlStr = sqlStr & " order by sorting asc" 
	rsget.Open sqlStr,dbget,1
 	if not rsget.eof then
 		fnGetEventSlideImg = rsget.getrows()
	end if
	rsget.close
End Function
 
 '//4단계 전체등록내용
 public Function fnGetEventST4
 	dim strSql
 	if FRectevtstate = "" THEN FRectevtstate = 0
 	if FRectevtstate > 5 then   
			strSql = "select e.evt_code,evt_kind,evt_manager,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,evt_lastupdate,adminid,evt_dispcate,dc1.catename as cateNm, dc2.catename as cateMNm" 	
		 	strSql = strSql & " ,issale,isgift,iscoupon,brand,evt_tag ,socname, etc_itemimg, evt_mo_listbanner ,mdtheme, title_pc, title_mo, evt_subcopyK,evt_subname,  saleper, salecper,themecolor,themecolormo,textbgcolor "
		 	strSql = strSql & " ,gift_isusing,gift_text1,gift_img1,gift_text2,gift_img2,gift_text3,gift_img3 "
		 	strSql = strSql & " from db_Event.dbo.tbl_event as e "
		 	strSql = strSql & "	inner join db_event.dbo.tbl_event_display as d on e.evt_code = d.evt_code "
		 	strSql = strSql & "	inner join db_event.[dbo].[tbl_event_md_theme] as m on d.evt_code = m.evt_code "
		 	strSql = strSql & " inner join  db_user.dbo.tbl_user_c as c on d.brand = c.userid  "
		 	strSql = strSql & "			left outer join db_item.dbo.tbl_display_cate as dc1 on   dc1.catecode = left(evt_dispcate,3) "&vbCrlf
		  strSql = strSql & " 			 left outer  join db_item.dbo.tbl_display_cate as dc2 on  dc2.catecode =  evt_dispcate  "&vbCrlf
		 	strSql= strSql & "  where e.evt_code = "&FevtCode&" and d.brand = '"&Fmakerid&"'" 	 
	else
			strSql = "select evt_code,evt_kind,evt_manager,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,evt_lastupdate,adminid,evt_dispcate,dc1.catename as cateNm, dc2.catename as cateMNm" 	
		 	strSql = strSql & " ,issale,isgift,iscoupon,brand,evt_tag ,socname, etc_itemimg, evt_mo_listbanner ,mdtheme, title_pc, title_mo, evt_subcopyK,evt_subname,  saleper, salecper,themecolor,themecolormo,textbgcolor "
		 	strSql = strSql & " ,gift_isusing,gift_text1,gift_img1,gift_text2,gift_img2,gift_text3,gift_img3 "
		 	strSql = strSql & " from db_Event.dbo.tbl_partner_event as e "
		 	strSql = strSql & " inner join  db_user.dbo.tbl_user_c as c on e.brand = c.userid  "
		 	strSql = strSql & "			left outer join db_item.dbo.tbl_display_cate as dc1 on   dc1.catecode = left(e.evt_dispcate,3) "&vbCrlf
		  	strSql = strSql & " 			 left outer  join db_item.dbo.tbl_display_cate as dc2 on  dc2.catecode =  e.evt_dispcate  "&vbCrlf
		 	strSql= strSql & "  where evt_code = "&FevtCode&" and brand = '"&Fmakerid&"'" 	
		 	 
	end if 
	 
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FevtCode 			= rsget("evt_code")
 	  Fevtkind     	 = rsget("evt_kind")       
 	  Fevtmanager   = rsget("evt_manager")    
 	  Fevtname       = rsget("evt_name")       
 	  Fevtstartdate  = rsget("evt_startdate")  
 	  Fevtenddate   = rsget("evt_enddate")    
 	  Fevtstate        = rsget("evt_state")      
 	  Fevtregdate     = rsget("evt_regdate")    
 	  Fevtusing        = rsget("evt_using")      
 	  Fevtlastupdate = rsget("evt_lastupdate") 
 	  Fadminid           = rsget("adminid")        
 	                                         
 	  Fevtdispcate  = rsget("evt_dispcate")    
 	  FevtCateNm		= rsget("catenm")   
 	  FevtCateMNm		= rsget("cateMnm")   
 	  Fissale             = rsget("issale")         
 	  Fisgift               = rsget("isgift")         
 	  Fiscoupon         = rsget("iscoupon")       
 	  Fbrand              = rsget("brand")          
 	  Fevttag            = rsget("evt_tag")       
 	 FbrandNm = rsget("socname")
 	  FTitlePC= rsget("title_pc")
 	  FTitleMO= rsget("title_mo")
 	  Fissale             = rsget("issale")          	  
 	  Fiscoupon         = rsget("iscoupon")    
 	  FsalePer 			= rsget("saleper")    
 	  FsaleCPer 			= rsget("salecper")    
	 Fetcitemimg          = rsget("etc_itemimg") 
	 Fevt_mo_listbanner  = rsget("evt_mo_listbanner")  
	 FsubcopyK           = rsget("evt_subcopyK")
	 Fevtsubname         = rsget("evt_subname")
	 Fmdtheme            = rsget("mdtheme")
	 Fthemecolor         = rsget("themecolor")
	 Fthemecolormo       = rsget("themecolormo")
	 Ftextbgcolor        = rsget("textbgcolor")
	 Fgiftisusing        = rsget("gift_isusing")
	 Fgifttext1          = rsget("gift_text1")
	 Fgiftimg1           = rsget("gift_img1")
	 Fgifttext2          = rsget("gift_text2")
	 Fgiftimg2           = rsget("gift_img2")
	 Fgifttext3          = rsget("gift_text3")
	 Fgiftimg3          = rsget("gift_img3")
	end if                               
	rsget.close                            
 End Function 
 
 public FLogevtCode 
 public Function fnGetEventLog
 dim strSql
  
 strSql = "select count(evtlog_code)"
 strSql = strSql & " FROM db_event.dbo.tbl_partner_eventStateLog "
 strSql = strSql & " where evt_code = "&FLogevtCode 
  rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FtotCnt = rsget(0)
end if
rsget.close
 strSql = "select evtlog_code, evt_code , evt_state, evt_text, filelink, l.regdate, evt_manager, regid , c.socname "
 strSql = strSql & " FROM db_event.dbo.tbl_partner_eventStateLog as l "
 strSql = strSql & "  left outer join db_user.dbo.tbl_user_c as c on l.regid = c.userid "
 strSql = strSql & " where evt_code = "&FLogevtCode&" order by evtlog_code desc " 	   
 
 rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventLog = rsget.getrows()
end if
rsget.close
End Function

public FCategoryPrdList()
public FItemArr
public FEItemCnt
public FItemsort

	Private Sub Class_Initialize()
		redim preserve FCategoryPrdList(0)
		FTotCnt = 0
		FItemArr = ""
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'##### 상품 리스트 ######
	public Function fnGetEventItem
		Dim strSql, arrItem,intI
		IF FEvtCode = "" THEN Exit Function
		IF FEvtGCode = "" THEN FEvtGCode= 0
		IF FEItemCnt ="" then FEItemCnt = 105
			IF FItemsort ="" then FItemsort =1
		strSql ="[db_item].[dbo].sp_Ten_partner_event_GetItem_new ("&FEvtCode&","&FEvtGCode&","&FEItemCnt&","&FItemsort&",1)"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrItem = rsget.GetRows()
		END IF
		rsget.close

		IF isArray(arrItem) THEN
			FTotCnt = Ubound(arrItem,2)
			redim preserve FCategoryPrdList(FTotCnt)

			For intI = 0 To FTotCnt
			set FCategoryPrdList(intI) = new CCategoryPrdItem
				FCategoryPrdList(intI).FItemID       = arrItem(0,intI)
				IF intI =0 THEN
				FItemArr = 	FCategoryPrdList(intI).FItemID
				ELSE
				FItemArr = FItemArr&","&FCategoryPrdList(intI).FItemID
				END IF
				FCategoryPrdList(intI).FItemName    = db2html(arrItem(1,intI))

				FCategoryPrdList(intI).FSellcash    = arrItem(2,intI)
				FCategoryPrdList(intI).FOrgPrice   	= arrItem(3,intI)
				FCategoryPrdList(intI).FMakerId   	= db2html(arrItem(4,intI))
				FCategoryPrdList(intI).FBrandName  	= db2html(arrItem(5,intI))

				FCategoryPrdList(intI).FSellYn      = arrItem(9,intI)
				FCategoryPrdList(intI).FSaleYn     	= arrItem(10,intI)
				FCategoryPrdList(intI).FLimitYn     = arrItem(11,intI)
				FCategoryPrdList(intI).FLimitNo     = arrItem(12,intI)
				FCategoryPrdList(intI).FLimitSold   = arrItem(13,intI)

				FCategoryPrdList(intI).FRegdate 		= arrItem(14,intI)
				FCategoryPrdList(intI).FReipgodate		= arrItem(15,intI)

                FCategoryPrdList(intI).Fitemcouponyn 	= arrItem(16,intI)
				FCategoryPrdList(intI).FItemCouponValue	= arrItem(17,intI)
				FCategoryPrdList(intI).Fitemcoupontype	= arrItem(18,intI)

				FCategoryPrdList(intI).Fevalcnt 		= arrItem(19,intI)
				FCategoryPrdList(intI).FitemScore 		= arrItem(20,intI)

				FCategoryPrdList(intI).FImageList		= "http://webimage.10x10.co.kr/image/list/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(6,intI)
				FCategoryPrdList(intI).FImageList120	= "http://webimage.10x10.co.kr/image/list120/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(7,intI)
				FCategoryPrdList(intI).FImageSmall		= "http://webimage.10x10.co.kr/image/small/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(8,intI)
				FCategoryPrdList(intI).FImageIcon1		= "http://webimage.10x10.co.kr/image/icon1/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(21,intI)
				FCategoryPrdList(intI).FImageIcon2		= "http://webimage.10x10.co.kr/image/icon2/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(22,intI)
				FCategoryPrdList(intI).FItemSize		= arrItem(23,intI)
				FCategoryPrdList(intI).Fitemdiv			= arrItem(24,intI)
				FCategoryPrdList(intI).FImageBasic		= "http://webimage.10x10.co.kr/image/basic/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(26,intI)
				FCategoryPrdList(intI).FImageBasic600	= "http://webimage.10x10.co.kr/image/basic/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(27,intI)
				FCategoryPrdList(intI).FfavCount		= arrItem(28,intI)

				If arrItem(29,intI) <> "" then
				FCategoryPrdList(intI).FAddImage		= "http://webimage.10x10.co.kr/image/add1/" & GetImageSubFolderByItemid(arrItem(0,intI)) & "/" & db2html(arrItem(29,intI))
				End if

				If Not(arrItem(31,intI)="" Or isnull(arrItem(31,intI))) Then 
					FCategoryPrdList(intI).Ftentenimage	= "http://webimage.10x10.co.kr/image/tenten/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(31,intI)
					FCategoryPrdList(intI).Ftentenimage50	= "http://webimage.10x10.co.kr/image/tenten50/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(32,intI)
					FCategoryPrdList(intI).Ftentenimage200	= "http://testwebimage.10x10.co.kr/image/tenten200/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(33,intI)
					FCategoryPrdList(intI).Ftentenimage400	= "http://testwebimage.10x10.co.kr/image/tenten400/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(34,intI)
					FCategoryPrdList(intI).Ftentenimage600	= "http://webimage.10x10.co.kr/image/tenten600/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(35,intI)
					FCategoryPrdList(intI).Ftentenimage1000	= "http://webimage.10x10.co.kr/image/tenten1000/"&GetImageSubFolderByItemid(arrItem(0,intI))&"/"&arrItem(36,intI)
				End If


			Next
		ELSE
			FTotCnt = -1
		END IF
	End Function
	
Sub sbSlidetemplateMD
	IF FevtCode = "" THEN Exit Sub	
	Dim vSArray , intSL , gubuncls		
	vSArray =fnGetEventSlideImg
 
		If isArray(vSArray) THEN 
			For intSL = 0 To UBound(vSArray,2)
	%>
	<div><img src="<%=vSArray(0,intSL)%>"></div>
	<%
			Next 
		End If 
	 
End Sub

Sub sbSlidetemplateItemMD 
	Dim arrImg, intLoop	  
	arrImg = fnGetEventItemImg
	If isArray(arrImg) Then 
		for intLoop=0 to ubound(arrImg,2)
	%>
	<div><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrImg(0,intLoop)%>"></div>
	<%
		Next 
	End If 
End Sub

	
Sub sbSlidetemplateMDMo
	IF FevtCode = "" THEN Exit Sub	
	Dim vSArray , intSL , gubuncls		
	vSArray =fnGetEventSlideImg
 
		If isArray(vSArray) THEN 
			For intSL = 0 To UBound(vSArray,2)
	%>
	<div class="swiper-slide"><div class="thumbnail"><img src="<%=vSArray(0,intSL)%>"></div></div>
	<%
			Next 
		End If 
	 
End Sub

Sub sbSlidetemplateItemMDMo 
	Dim arrImg, intLoop	  
	arrImg = fnGetEventItemImg
	If isArray(arrImg) Then 
		for intLoop=0 to ubound(arrImg,2)
	%>
	<div class="swiper-slide"><div class="thumbnail"><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrImg(0,intLoop)%>"></div></div>
	<%
		Next 
	End If 
End Sub

Sub sbEvtItemView
	Dim intIx, sBadges,itemid

	IF eCode = "" THEN Exit Sub
	intI = 0

	call fnGetEventItem
	iTotCnt = FTotCnt

	IF itemid = "" THEN
		itemid = FItemArr
	ELSE
		itemid = itemid&","&FItemArr
	END If
	
	intI = 0 
		IF (iTotCnt >= 0) THEN
		if FCategoryPrdList(0).FItemSize="2" or FCategoryPrdList(0).FItemSize="200" Then
			IF blnItemifno THEN 
			%>
			<div class="pdtWrap pdt400V15">
				<ul class="pdtList">
			<%
				For intI =0 To iTotCnt
					'큰이미지가 끝나면 출력 종료
					if FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="150" or FCategoryPrdList(intI).FItemSize="153" or FCategoryPrdList(intI).FItemSize="155" or FCategoryPrdList(intI).FItemSize="160" then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
					<div class="pdtBox">
						<div class="pdtPhoto">
							<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage400="" Or isnull(FCategoryPrdList(intI).Ftentenimage400)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage400%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"400","400","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"400","400","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
						</div>
						<div class="pdtInfo">
							<p class="pdtBrand tPad20"><a href="/street/street_brand.asp?makerid=<%=FCategoryPrdList(intI).FMakerId %>"><%=FCategoryPrdList(intI).FBrandName %></a></p>
							<p class="pdtName tPad07"><a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><%=FCategoryPrdList(intI).FItemName%></a></p>
							<% If blnitempriceyn = "1" Then %>
							<% Else %>
								<% if FCategoryPrdList(intI).IsSaleItem or FCategoryPrdList(intI).isCouponItem Then %>
									<% IF FCategoryPrdList(intI).IsSaleItem then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0)%>원</span> <strong class="cRd0V15">[<%=FCategoryPrdList(intI).getSalePro%>]</strong></p>
									<% End If %>
									<% IF FCategoryPrdList(intI).IsCouponItem Then %>
										<% if Not(FCategoryPrdList(intI).IsFreeBeasongCoupon() or FCategoryPrdList(intI).IsSaleItem) Then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
										<% end If %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).GetCouponAssignPrice,0)%>원</span> <strong class="cGr0V15">[<%=FCategoryPrdList(intI).GetCouponDiscountStr%>]</strong></p>
									<% End If %>
								<% Else %>
								<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0) & chkIIF(FCategoryPrdList(intI).IsMileShopitem,"Point","원")%></span></p>
								<% End If %>
							<p class="pdtStTag tPad10">
								<% IF FCategoryPrdList(intI).isSoldOut Then %>
									<img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" />
								<% else %>
									<% IF FCategoryPrdList(intI).isTempSoldOut Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" /> <% end if %>
									<% IF FCategoryPrdList(intI).isSaleItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_sale.gif" alt="SALE" /> <% end if %>
									<% IF FCategoryPrdList(intI).isCouponItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_coupon.gif" alt="쿠폰" /> <% end if %>
									<% IF FCategoryPrdList(intI).isLimitItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_limited.gif" alt="한정" /> <% end if %>
									<% IF FCategoryPrdList(intI).IsTenOnlyitem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_only.gif" alt="ONLY" /> <% end if %>
									<% IF FCategoryPrdList(intI).isNewItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_new.gif" alt="NEW" /> <% end if %>
								<% end if %>
							</p>
							<% End If %>
						</div>
						<ul class="pdtActionV15">
							<li class="largeView"><a href="" onclick="ZoomItemInfo('<%=FCategoryPrdList(intI).FItemid %>'); return false;"><img src="http://fiximage.10x10.co.kr/web2015/common/btn_quick.png" alt="QUICK" /></a></li>
							<li class="postView"><a href="" <%=chkIIF(FCategoryPrdList(intI).Fevalcnt>0,"onclick=""popEvaluate('" & FCategoryPrdList(intI).FItemid & "');return false;""","onclick=""return false;""")%>><span><%=FCategoryPrdList(intI).Fevalcnt%></span></a></li>
							<li class="wishView"><a href="" onclick="TnAddFavorite('<%=FCategoryPrdList(intI).FItemid %>');return false;"><span><%=FCategoryPrdList(intI).FfavCount%></span></a></li>
						</ul>
					</div>
					</li>
			<%
					set FCategoryPrdList(intI) = nothing
				Next
			%>
				</ul>
			</div>
		 <% Else %>
			<div class="pdtWrap pdt400V15">
				<ul class="pdtList">
			<%
				For intI =0 To iTotCnt
					if FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="150" or FCategoryPrdList(intI).FItemSize="153" or FCategoryPrdList(intI).FItemSize="155" or FCategoryPrdList(intI).FItemSize="160" then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
						<div class="pdtBox">
							<div class="pdtPhoto">
								<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage400="" Or isnull(FCategoryPrdList(intI).Ftentenimage400)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage400%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"400","400","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"400","400","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
							</div>
						</div>
					</li>
			<%
					Set FCategoryPrdList(intI) = nothing
				Next 
			%>
				</ul>
			</div>
<%
			End If
		end if
	end If
	'// 이미지 사이즈가 중간일경우(240px:4개) 표시(2017-08-07; 정태훈) 추가
	intIx = intI

	IF (iTotCnt >= intIx) THEN
		if FCategoryPrdList(intI).FItemSize="153" then
			IF blnItemifno THEN 
			%>
			<div class="pdtWrap pdt240V15">
				<ul class="pdtList">
			<%
				For intI = intIx To iTotCnt
			
					'중간이미지가 끝나면 출력 종료
					if FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="150" then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
					<div class="pdtBox">
						<div class="pdtPhoto">
							<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage400="" Or isnull(FCategoryPrdList(intI).Ftentenimage400)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage400%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"240","240","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"240","240","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
						</div>
						<div class="pdtInfo">
							<p class="pdtBrand tPad20"><a href="/street/street_brand.asp?makerid=<%=FCategoryPrdList(intI).FMakerId %>"><%=FCategoryPrdList(intI).FBrandName %></a></p>
							<p class="pdtName tPad07"><a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><%=FCategoryPrdList(intI).FItemName%></a></p>
							<% If blnitempriceyn = "1" Then %>
							<% Else %>
								<% if FCategoryPrdList(intI).IsSaleItem or FCategoryPrdList(intI).isCouponItem Then %>
									<% IF FCategoryPrdList(intI).IsSaleItem then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0)%>원</span> <strong class="cRd0V15">[<%=FCategoryPrdList(intI).getSalePro%>]</strong></p>
									<% End If %>
									<% IF FCategoryPrdList(intI).IsCouponItem Then %>
										<% if Not(FCategoryPrdList(intI).IsFreeBeasongCoupon() or FCategoryPrdList(intI).IsSaleItem) Then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
										<% end If %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).GetCouponAssignPrice,0)%>원</span> <strong class="cGr0V15">[<%=FCategoryPrdList(intI).GetCouponDiscountStr%>]</strong></p>
									<% End If %>
								<% Else %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0) & chkIIF(FCategoryPrdList(intI).IsMileShopitem,"Point","원")%></span></p>
								<% End If %>
							<p class="pdtStTag tPad10">
								<% IF FCategoryPrdList(intI).isSoldOut Then %>
									<img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" />
								<% else %>
									<% IF FCategoryPrdList(intI).isTempSoldOut Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" /> <% end if %>
									<% IF FCategoryPrdList(intI).isSaleItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_sale.gif" alt="SALE" /> <% end if %>
									<% IF FCategoryPrdList(intI).isCouponItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_coupon.gif" alt="쿠폰" /> <% end if %>
									<% IF FCategoryPrdList(intI).isLimitItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_limited.gif" alt="한정" /> <% end if %>
									<% IF FCategoryPrdList(intI).IsTenOnlyitem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_only.gif" alt="ONLY" /> <% end if %>
									<% IF FCategoryPrdList(intI).isNewItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_new.gif" alt="NEW" /> <% end if %>
								<% end if %>
							</p>
							<% End If %>
						</div>
						<ul class="pdtActionV15">
							<li class="largeView"><a href="" onclick="ZoomItemInfo('<%=FCategoryPrdList(intI).FItemid %>'); return false;"><img src="http://fiximage.10x10.co.kr/web2015/common/btn_quick.png" alt="QUICK" /></a></li>
							<li class="postView"><a href="" <%=chkIIF(FCategoryPrdList(intI).Fevalcnt>0,"onclick=""popEvaluate('" & FCategoryPrdList(intI).FItemid & "');return false;""","onclick=""return false;""")%>><span><%=FormatNumber(FCategoryPrdList(intI).Fevalcnt,0)%></span></a></li>
							<li class="wishView"><a href="" onclick="TnAddFavorite('<%=FCategoryPrdList(intI).FItemid %>');return false;"><span><%=FormatNumber(FCategoryPrdList(intI).FfavCount,0)%></span></a></li>
						</ul>
					</div>
					</li>
			<%
					set FCategoryPrdList(intI) = nothing
				Next
			%>
				</ul>
			</div>
		 <% Else %>
			<div class="pdtWrap pdt240V15">
				<ul class="pdtList">
			<%
				For intI =intIx To iTotCnt
					if FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="150" then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
						<div class="pdtBox">
							<div class="pdtPhoto">
								<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage400="" Or isnull(FCategoryPrdList(intI).Ftentenimage400)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage400%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"240","240","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"240","240","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
							</div>
						</div>
					</li>
			<%
					Set FCategoryPrdList(intI) = nothing
				Next 
			%>
				</ul>
			</div>
<%
			End If
		end if
	end If
	'// 이미지 사이즈가 중간일경우(200px 기존 -> 180xp 변경)
	intIx = intI

	IF (iTotCnt >= intIx) THEN
		if FCategoryPrdList(intI).FItemSize="150" then
			IF blnItemifno THEN 
			%>
			<div class="pdtWrap pdt180V15">
				<ul class="pdtList">
			<%
				For intI = intIx To iTotCnt
			
					'중간이미지가 끝나면 출력 종료
					if FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="100" then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
					<div class="pdtBox">
						<div class="pdtPhoto">
							<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage200="" Or isnull(FCategoryPrdList(intI).Ftentenimage200)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage200%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"180","180","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"180","180","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
						</div>
						<div class="pdtInfo">
							<p class="pdtBrand tPad20"><a href="/street/street_brand.asp?makerid=<%=FCategoryPrdList(intI).FMakerId %>"><%=FCategoryPrdList(intI).FBrandName %></a></p>
							<p class="pdtName tPad07"><a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><%=FCategoryPrdList(intI).FItemName%></a></p>
							<% If blnitempriceyn = "1" Then %>
							<% Else %>
								<% if FCategoryPrdList(intI).IsSaleItem or FCategoryPrdList(intI).isCouponItem Then %>
									<% IF FCategoryPrdList(intI).IsSaleItem then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0)%>원</span> <strong class="cRd0V15">[<%=FCategoryPrdList(intI).getSalePro%>]</strong></p>
									<% End If %>
									<% IF FCategoryPrdList(intI).IsCouponItem Then %>
										<% if Not(FCategoryPrdList(intI).IsFreeBeasongCoupon() or FCategoryPrdList(intI).IsSaleItem) Then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
										<% end If %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).GetCouponAssignPrice,0)%>원</span> <strong class="cGr0V15">[<%=FCategoryPrdList(intI).GetCouponDiscountStr%>]</strong></p>
									<% End If %>
								<% Else %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0) & chkIIF(FCategoryPrdList(intI).IsMileShopitem,"Point","원")%></span></p>
								<% End If %>
							<p class="pdtStTag tPad10">
								<% IF FCategoryPrdList(intI).isSoldOut Then %>
									<img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" />
								<% else %>
									<% IF FCategoryPrdList(intI).isTempSoldOut Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" /> <% end if %>
									<% IF FCategoryPrdList(intI).isSaleItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_sale.gif" alt="SALE" /> <% end if %>
									<% IF FCategoryPrdList(intI).isCouponItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_coupon.gif" alt="쿠폰" /> <% end if %>
									<% IF FCategoryPrdList(intI).isLimitItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_limited.gif" alt="한정" /> <% end if %>
									<% IF FCategoryPrdList(intI).IsTenOnlyitem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_only.gif" alt="ONLY" /> <% end if %>
									<% IF FCategoryPrdList(intI).isNewItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_new.gif" alt="NEW" /> <% end if %>
								<% end if %>
							</p>
							<% End If %>
						</div>
						<ul class="pdtActionV15">
							<li class="largeView"><a href="" onclick="ZoomItemInfo('<%=FCategoryPrdList(intI).FItemid %>'); return false;"><img src="http://fiximage.10x10.co.kr/web2015/common/btn_quick.png" alt="QUICK" /></a></li>
							<li class="postView"><a href="" <%=chkIIF(FCategoryPrdList(intI).Fevalcnt>0,"onclick=""popEvaluate('" & FCategoryPrdList(intI).FItemid & "');return false;""","onclick=""return false;""")%>><span><%=FormatNumber(FCategoryPrdList(intI).Fevalcnt,0)%></span></a></li>
							<li class="wishView"><a href="" onclick="TnAddFavorite('<%=FCategoryPrdList(intI).FItemid %>');return false;"><span><%=FormatNumber(FCategoryPrdList(intI).FfavCount,0)%></span></a></li>
						</ul>
					</div>
					</li>
			<%
					set FCategoryPrdList(intI) = nothing
				Next
			%>
				</ul>
			</div>
		 <% Else %>
			<div class="pdtWrap pdt200V15">
				<ul class="pdtList">
			<%
				For intI =intIx To iTotCnt
					If FCategoryPrdList(intI).FItemSize="1" or FCategoryPrdList(intI).FItemSize="100" Then Exit For
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
						<div class="pdtBox">
							<div class="pdtPhoto">
								<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><% if Not(FCategoryPrdList(intI).Ftentenimage200="" Or isnull(FCategoryPrdList(intI).Ftentenimage200)) Then %> <img src="<%=FCategoryPrdList(intI).Ftentenimage200%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% Else %> <img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"200","200","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /> <% End If %> <% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"200","200","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
							</div>
						</div>
					</li>
			<%
					Set FCategoryPrdList(intI) = nothing
				Next 
			%>
				</ul>
			</div>
<%
			End If
		end if
	end if

	'// 일반 사이즈 상품 목록 출력
	intIx = intI

	IF (iTotCnt >= intIx) THEN
		IF blnItemifno THEN 
%>
			<div class="pdtWrap pdt130V15">
				<ul class="pdtList">
			<%
				For intI =intIx To iTotCnt
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
					<div class="pdtBox">
						<div class="pdtPhoto">
							<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"130","130","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /><% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"130","130","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
						</div>
						<div class="pdtInfo">
							<p class="pdtBrand tPad20"><a href="/street/street_brand.asp?makerid=<%=FCategoryPrdList(intI).FMakerId %>"><%=FCategoryPrdList(intI).FBrandName %></a></p>
							<p class="pdtName tPad07"><a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><%=FCategoryPrdList(intI).FItemName%></a></p>
							<% If blnitempriceyn = "1" Then %>
							<% Else %>
								<% if FCategoryPrdList(intI).IsSaleItem or FCategoryPrdList(intI).isCouponItem Then %>
									<% IF FCategoryPrdList(intI).IsSaleItem then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0)%>원</span> <strong class="cRd0V15">[<%=FCategoryPrdList(intI).getSalePro%>]</strong></p>
									<% End If %>
									<% IF FCategoryPrdList(intI).IsCouponItem Then %>
										<% if Not(FCategoryPrdList(intI).IsFreeBeasongCoupon() or FCategoryPrdList(intI).IsSaleItem) Then %>
									<p class="pdtPrice"><span class="txtML"><%=FormatNumber(FCategoryPrdList(intI).getOrgPrice,0)%>원</span></p>
										<% end If %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).GetCouponAssignPrice,0)%>원</span> <strong class="cGr0V15">[<%=FCategoryPrdList(intI).GetCouponDiscountStr%>]</strong></p>
									<% End If %>
								<% Else %>
									<p class="pdtPrice"><span class="finalP"><%=FormatNumber(FCategoryPrdList(intI).getRealPrice,0) & chkIIF(FCategoryPrdList(intI).IsMileShopitem,"Point","원")%></span></p>
								<% End If %>
							<p class="pdtStTag tPad10">
								<% IF FCategoryPrdList(intI).isSoldOut Then %>
									<img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" />
								<% else %>
									<% IF FCategoryPrdList(intI).isTempSoldOut Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_soldout.gif" alt="SOLDOUT" /> <% end if %>
									<% IF FCategoryPrdList(intI).isSaleItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_sale.gif" alt="SALE" /> <% end if %>
									<% IF FCategoryPrdList(intI).isCouponItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_coupon.gif" alt="쿠폰" /> <% end if %>
									<% IF FCategoryPrdList(intI).isLimitItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_limited.gif" alt="한정" /> <% end if %>
									<% IF FCategoryPrdList(intI).IsTenOnlyitem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_only.gif" alt="ONLY" /> <% end if %>
									<% IF FCategoryPrdList(intI).isNewItem Then %><img src="http://fiximage.10x10.co.kr/web2013/shopping/tag_new.gif" alt="NEW" /> <% end if %>
								<% end if %>
							</p>
							<% End If %>
						</div>
						<ul class="pdtActionV15">
							<li class="largeView"><a href="" onclick="ZoomItemInfo('<%=FCategoryPrdList(intI).FItemid %>'); return false;"><img src="http://fiximage.10x10.co.kr/web2015/common/btn_quick.png" alt="QUICK" /></a></li>
							<li class="postView"><a href="" <%=chkIIF(FCategoryPrdList(intI).Fevalcnt>0,"onclick=""popEvaluate('" & FCategoryPrdList(intI).FItemid & "');return false;""","onclick=""return false;""")%>><span><%=FormatNumber(FCategoryPrdList(intI).Fevalcnt,0)%></span></a></li>
							<li class="wishView"><a href="" onclick="TnAddFavorite('<%=FCategoryPrdList(intI).FItemid %>');return false;"><span><%=FormatNumber(FCategoryPrdList(intI).FfavCount,0)%></span></a></li>
						</ul>
					</div>
					</li>
			<%
					set FCategoryPrdList(intI) = nothing
				Next
			%>
				</ul>
			</div>
			<%set cEventItem = nothing%>
	   <% Else %>
			<div class="pdtWrap pdt130V15">
				<ul class="pdtList">
			<%
				For intI =intIx To iTotCnt
			%>
					<li <%=chkIIF(FCategoryPrdList(intI).isSoldOut," class=""soldOut""","")%>>
						<div class="pdtBox">
							<div class="pdtPhoto">
								<a href="/shopping/category_prd.asp?itemid=<%=FCategoryPrdList(intI).FItemID %>"><span class="soldOutMask"></span><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FImageBasic,"130","130","true","false")%>" alt="<%=FCategoryPrdList(intI).FItemName%>" /><% if FCategoryPrdList(intI).FAddimage<>"" then %><dfn><img src="<%=getThumbImgFromURL(FCategoryPrdList(intI).FAddimage,"130","130","true","false")%>" alt="<%=Replace(FCategoryPrdList(intI).FItemName,"""","")%>" /></dfn><% end if %></a>
							</div>
						</div>
					</li>
			<%
					Set FCategoryPrdList(intI) = nothing
				Next 
			%>
				</ul>
			</div>
<%			
			Set cEventItem = nothing
		End If
	End IF
End Sub



'----------------------------------------------------
' sbEvtItemView_2015 : 상품목록 보여주기 (일반형태)
' 2015-06-19 ver2.0 이종화
'----------------------------------------------------
Sub sbEvtItemView_2015
	Dim iEndCnt, intJ, iLp, vWishArr
dim itemid,eItemListType
eItemListType = 1
	IF eCode = "" THEN Exit Sub
	intI = 0

 
	 call fnGetEventItem
	iTotCnt = FTotCnt

	IF itemid = "" THEN
		itemid = FItemArr
	ELSE
		itemid = itemid&","&FItemArr
	END IF


	IF (iTotCnt >= 0) THEN
		If eItemListType = "1" Then '### 격자형
			Response.Write "				<div class=""items type-grid"">"
		ElseIf eItemListType = "2" Then '### 리스트형
			Response.Write "				<div class=""items type-list"">"
		ElseIf eItemListType = "3" Then '### BIG형
			Response.Write "				<div class=""items type-big"">"
		End If
%>
		<!--<div class="pdtListWrapV15a">-->
		<ul>
<%
			For intI =0 To iTotCnt
%>
				<li>
					<a href="/category/category_itemPrd.asp?itemid=<% =FCategoryPrdList(intI).Fitemid %>">
						<!-- for dev msg : 상품명으로 썸네일 alt값 달면 중복되니 alt=""으로 처리해주세요. -->
						<div class="thumbnail">
							<% if (eItemListType = "1") or (eItemListType = "2") then %>
								<img src="<%=chkiif(Not( FCategoryPrdList(intI).Ftentenimage400="" Or isnull( FCategoryPrdList(intI).Ftentenimage400)), FCategoryPrdList(intI).Ftentenimage400,getThumbImgFromURL( FCategoryPrdList(intI).FImageBasic,300,300,"true","false")) %>" alt="<% = FCategoryPrdList(intI).FItemName %>" />
							<% else %>
								<img src="<%=chkiif(Not( FCategoryPrdList(intI).Ftentenimage400="" Or isnull( FCategoryPrdList(intI).Ftentenimage400)), FCategoryPrdList(intI).Ftentenimage400,getThumbImgFromURL( FCategoryPrdList(intI).FImageBasic,400,400,"true","false")) %>" alt="<% = FCategoryPrdList(intI).FItemName %>" />
							<% end if %>
							<% IF FCategoryPrdList(intI).IsSoldOut Or FCategoryPrdList(intI).isTempSoldOut Then %>
								<b class="soldout">일시 품절</b>
							<% end if %>
						</div>
						<div class="desc">
							<span class="brand"><% = FCategoryPrdList(intI).FBrandName %></span>
							<p class="name"><% = FCategoryPrdList(intI).FItemName %></p>
							<div class="price">
								<%
									If FCategoryPrdList(intI).IsSaleItem AND FCategoryPrdList(intI).isCouponItem Then	'### 쿠폰 O 세일 O
										Response.Write "<div class=""unit""><b class=""sum"">" & FormatNumber( FCategoryPrdList(intI).GetCouponAssignPrice,0) & "<span class=""won"">원</span></b>"
										Response.Write "&nbsp;<b class=""discount color-red"">" & FCategoryPrdList(intI).getSalePro & "</b>"
										If FCategoryPrdList(intI).Fitemcoupontype <> "3" Then	'### 무료배송아닌것
											If InStr( FCategoryPrdList(intI).GetCouponDiscountStr,"%") < 1 Then	'### 금액 쿠폰은 쿠폰으로 표시
												Response.Write "&nbsp;<b class=""discount color-green""><small>쿠폰</small></b>"
											Else
												Response.Write "&nbsp;<b class=""discount color-green"">" & FCategoryPrdList(intI).GetCouponDiscountStr & "<small>쿠폰</small></b>"
											End If
										End If
										Response.Write "</div>" &  vbCrLf
									ElseIf FCategoryPrdList(intI).IsSaleItem AND (Not FCategoryPrdList(intI).isCouponItem) Then	'### 쿠폰 X 세일 O
										Response.Write "<div class=""unit""><b class=""sum"">" & FormatNumber( FCategoryPrdList(intI).getRealPrice,0) & "<span class=""won"">원</span></b>"
										Response.Write "&nbsp;<b class=""discount color-red"">" & FCategoryPrdList(intI).getSalePro & "</b>"
										Response.Write "</div>" &  vbCrLf
									ElseIf FCategoryPrdList(intI).isCouponItem AND (NOT FCategoryPrdList(intI).IsSaleItem) Then	'### 쿠폰 O 세일 X
										Response.Write "<div class=""unit""><b class=""sum"">" & FormatNumber( FCategoryPrdList(intI).GetCouponAssignPrice,0) & "<span class=""won"">원</span></b>"
										If FCategoryPrdList(intI).Fitemcoupontype <> "3" Then	'### 무료배송아닌것
											If InStr( FCategoryPrdList(intI).GetCouponDiscountStr,"%") < 1 Then	'### 금액 쿠폰은 쿠폰으로 표시
												Response.Write "&nbsp;<b class=""discount color-green""><small>쿠폰</small></b>"
											Else
												Response.Write "&nbsp;<b class=""discount color-green"">" & FCategoryPrdList(intI).GetCouponDiscountStr & "<small>쿠폰</small></b>"
											End If
										End If
										Response.Write "</div>" &  vbCrLf
									Else
										Response.Write "<div class=""unit""><b class=""sum"">" & FormatNumber( FCategoryPrdList(intI).getRealPrice,0) & "<span class=""won"">" & CHKIIF( FCategoryPrdList(intI).IsMileShopitem," Point","원") & "</span></b></div>" &  vbCrLf
									End If
								%>
							</div>
						</div>
					</a>
					<div class="etc">
						<!-- for dev msg : 리뷰
							1. 리뷰수와 wish수가 1,000건 이상이면 999+로 표시해주세요
							2. 리뷰는 총 평점으로 퍼센트로 표현해주세요. <i style="width:50%;">...</i>
						--> 
						
						<button class="tag wish btn-wish" onclick="goWishPop('<%= FCategoryPrdList(intI).FItemid%>','<%=eCode%>');">
						</button>
						<% IF FCategoryPrdList(intI).IsCouponItem AND FCategoryPrdList(intI).GetCouponDiscountStr = "무료배송" Then %>
							<div class="tag shipping"><span class="icon icon-shipping"><i>무료배송</i></span> FREE</div>
						<% End If %>
					</div>
				</li>
<%
			Next
		Response.write "</ul>"	'</div>
		Response.Write "</div>"
	End IF
End Sub
End Class                                
%>

 