<%

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
public  FbrandNm 
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
  
public FRectItemid
public FRectItemName
public FPageSize
public FCurrPage
public FRectSailYn
public FRectSort
public FRectSellYN 
public FRectCouponYn

public FitemTotCnt
 public FTotCnt
 public FPSize
 public FCPage
 public FSPageNo
 public FEPageNo
 
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
 
  
  if FRectState <> "" and 	FRectState <> "A" then  
  	strSearch = strSearch & " and evt_state in ("&FRectState&")"
  end if

if FRectDispcate <> "" then
  	strSearch = strSearch & "  and left(evt_dispcate,"&len(FRectDispcate)&")="&FRectDispcate
end if


  if FRectNm <> "" then
  	strSearch = strSearch & "  and evt_name like '%"&FRectNm&"%'"
end if
  
  if FRectmakerid <> "" then
  	  	strSearch = strSearch & "  and e.brand ='"&FRectmakerid&"'" 
end if

  strSql = " SELECT count(evt_code)  "
  strSql = strSql & " 		FROM db_event.[dbo].[tbl_partner_event] as e "&vbCrlf
  strSql = strSql & "			left outer join db_item.dbo.tbl_display_cate as dc1 on left(e.evt_dispcate,3) = dc1.catecode "&vbCrlf
  strSql = strSql & " 			 left outer join db_item.dbo.tbl_display_cate as dc2 on e.evt_dispcate = dc2.catecode "&vbCrlf
  strSql = strSql &"   	where e.evt_using = 'Y' and e.evt_state > 0 " & strSearch   
 
  rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
	rsget.close

if 		FTotCnt > 0 then	
	FSPageNo = (FPSize*(FCPage-1)) + 1
	FEPageNo = FPSize*FCPage	
	
	
  strSql = "SELECT  evt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,adminid,evt_dispcate,brand,salePer,saleCPer,mdtheme ,dc1nm, dc2nm "&vbCrlf
  strSql = strSql & " FROM ( "&vbCrlf
  strSql = strSql & "  	SELECT ROW_NUMBER() OVER (ORDER BY evt_code desc ) as RowNum "&vbCrlf
  strSql = strSql & "			,evt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,adminid,evt_dispcate,brand,salePer,saleCPer,mdtheme ,dc1.catename as dc1nm, dc2.catename as dc2nm "&vbCrlf
  strSql = strSql & " 		FROM db_event.[dbo].[tbl_partner_event] as e "&vbCrlf
  strSql = strSql & "			left outer join db_item.dbo.tbl_display_cate as dc1 on   dc1.catecode =left(e.evt_dispcate,3) "&vbCrlf
  strSql = strSql & " 			 left outer join db_item.dbo.tbl_display_cate as dc2 on  dc2.catecode = e.evt_dispcate  "&vbCrlf
  strSql = strSql &"   	where e.evt_using = 'Y'  and e.evt_state > 0 " & strSearch&vbCrlf
  strSql = strSql & " ) as TB "&vbCrlf
  strSql = strSql & " WHERE TB.RowNum Between "&FSPageNo&" AND "  &FEPageNo &vbCrlf
 strSql = strSql & " order by evt_code desc "
   rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventList = rsget.getRows()
	end if
	rsget.close
end if	
End Function

 


'//2단계 그룹리스트
public Function fnGetEventGroup
 	dim strSql
 	strSql = "select g.evtgroup_code, g.evtgroup_desc, g.evtgroup_sort "
 	strSql = strSql & ",(select count(itemid) from db_event.dbo.tbl_partner_eventitem where evt_code = g.evt_code and evtgroup_code = g.evtgroup_code and evtitem_isUsing =1) as evtitem_cnt"
 	strSql = strSql & " from db_event.dbo.tbl_partner_eventitem_group as g "
 	strSql = strSql & "	inner join db_event.dbo.tbl_partner_event as e on g.evt_code = e.evt_code  "
 	strSql = strSql & " where g.evt_code="&FevtCode&" and g.evtgroup_using ='Y' and g.evtgroup_pcode >0 order by g.evtgroup_sort "
 	 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventGroup = rsget.getRows()
	end if
	rsget.close
End function


'//2단계 그룹 등록상품 리스트
public Function fnGetEventGroupItem
	dim strSql
	strSql ="SELECT count(itemid) FROM db_event.[dbo].[tbl_partner_eventitem] where evt_code ="&FevtCode&" and evtgroup_code ="&FevtGCode& " and evtitem_isusing = 1 "
		rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
 	rsget.close
 	
 	if FTotCnt >0 then
	strSql = "SELECT  ei.[itemid],i.itemname, ei.[evtitem_sort],ei.[evtitem_imgsize],sailyn, sellcash, buycash, orgprice, orgsuplycash, sailprice, sailsuplycash,   itemcouponyn, itemcoupontype, itemcouponvalue, sellyn  "
  	strSql = strSql & " FROM db_event.[dbo].[tbl_partner_eventitem] as ei "  	
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



                      
public Function fnGetItemtMaxSale
	dim strSql 
	if FevtCode ="" then exit Function
	if FRectSailYn ="Y" then
		strSql ="  select max(((orgprice-sailprice)/ orgprice)*100) as sailpercent  "
		strSql = strSql & " FROM db_event.[dbo].[tbl_partner_eventitem] as ei  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item as i on ei.itemid = i.itemid  "
		strSql = strSql & " WHERE ei.evt_code="&FevtCode&"   and evtitem_isusing = 1 and i.sailyn='Y'" 
	elseif FRectCouponYn="Y" then
		strSql ="  select  itemcouponvalue  "
		strSql = strSql & " FROM db_event.[dbo].[tbl_partner_eventitem] as ei  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item as i on ei.itemid = i.itemid  "
		strSql = strSql & " WHERE ei.evt_code="&FevtCode&"   and evtitem_isusing = 1 and i.itemCouponyn='Y' and itemcoupontype =1 " 
	end if	
	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetItemtMaxSale = rsget(0)
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
 dim strSql
 strSql = "select count(b.idx) "
 strSql = strSql & " from [db_event].[dbo].[tbl_partner_event_itembanner] as b "
 strSql = strSql & " inner join db_item.dbo.tbl_item as i on b.itemid = i.itemid "
 strSql = strSql & " where b.evt_code = "&FevtCode&"  and b.sdiv='"&FSdiv&"'" 	 
 rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FTotCnt = rsget(0)
 	end if
 	rsget.close
 	
 strSql = "select b.idx, b.itemid, b.viewidx, i.itemname,sailyn, sellcash, buycash, orgprice, orgsuplycash, sailprice, sailsuplycash,   itemcouponyn, itemcoupontype, itemcouponvalue"
 strSql = strSql & " from [db_event].[dbo].[tbl_partner_event_itembanner] as b "
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
dim sqlStr
	sqlStr = "select  top 3  i.basicimage, i.itemid ,i.listimage "
	sqlStr = sqlStr & " from [db_event].[dbo].[tbl_partner_event_itembanner] e "
	sqlStr = sqlStr & "	join [db_item].[dbo].tbl_item i on e.itemid = i.itemid "
	sqlStr= sqlStr &" where e.evt_code= '" & FevtCode & "'   and e.sdiv='"&FSdiv&"' "
	sqlStr = sqlStr & " order by e.viewidx asc"
	rsget.Open sqlStr,dbget,1
 	if not rsget.eof then
 		fnGetEventItemImg = rsget.getrows()
	end if
	rsget.close
End Function

 
  
 '//3단계 이미지
 public Function fnGetEventSlideImg
dim sqlStr
	sqlStr = "select  top 3 slideimg, idx, device  "
	sqlStr = sqlStr & " from [db_event].[dbo].[tbl_partner_event_slide_addimage]  "
	sqlStr= sqlStr &" where evt_code= '" & FevtCode & "'   and device='"&FSdiv&"' and isusing ='Y' "
	sqlStr = sqlStr & " order by sorting asc"
	rsget.Open sqlStr,dbget,1
 	if not rsget.eof then
 		fnGetEventSlideImg = rsget.getrows()
	end if
	rsget.close
End Function

 public Fevtdispcate
 public FevtText
  public Ffilelink
 '//4단계 전체등록내용
 public Function fnGetEventST4
 	dim strSql
 	strSql = "select e.evt_code,evt_kind,evt_manager,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,evt_lastupdate,adminid,evt_dispcate,dc1.catename as cateNm, dc2.catename as cateMNm" 	&vbCrlf
 	strSql = strSql & " ,issale,isgift,iscoupon,brand,evt_tag ,socname, etc_itemimg, evt_mo_listbanner ,mdtheme, title_pc, title_mo, evt_subcopyK,evt_subname,  saleper, salecper,themecolor,themecolormo,textbgcolor "&vbCrlf
 	strSql = strSql & " ,gift_isusing,gift_text1,gift_img1,gift_text2,gift_img2,gift_text3,gift_img3 "&vbCrlf
 	strSql = strSql & " , l.evt_text , l.filelink "&vbCrlf
 	strSql = strSql & " from db_Event.dbo.tbl_partner_event as e "&vbCrlf
 	strSql = strSql & " inner join  db_user.dbo.tbl_user_c as c on e.brand = c.userid  "&vbCrlf
 	strSql = strSql & "	left outer join db_item.dbo.tbl_display_cate as dc1 on   dc1.catecode = left(e.evt_dispcate,3)  "&vbCrlf
  strSql = strSql & " left outer join db_item.dbo.tbl_display_cate as dc2 on  dc2.catecode =  e.evt_dispcate  "&vbCrlf 
  strSql = strSql & " left outer join ( select top 1 evt_code, evt_text, filelink from db_event.dbo.tbl_partner_eventStateLog where evt_code =  "&FevtCode &" and evt_state = 5  order by evtlog_code desc ) as l on l.evt_code = e.evt_code " 
 	strSql= strSql & "  where e.evt_code = "&FevtCode  
 
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
	 FevtText 					= rsget("evt_text")
	 Ffilelink 					= rsget("filelink")
	end if                               
	rsget.close                            
 End Function 
  
 public Function fnGetEventLog
 dim strSql
 strSql = "select count(evtlog_code)"
 strSql = strSql & " FROM db_event.dbo.tbl_partner_eventStateLog "
 strSql = strSql & " where evt_code = "&FevtCode 
  rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		FtotCnt = rsget(0)
end if
rsget.close
 strSql = "select evtlog_code, evt_code , evt_state, evt_text, filelink, l.regdate, evt_manager, regid , "
 strSql = strSql &" case evt_manager when '1' then t.username "
  strSql= strSql &" when '2' then c.socname end as regNm"
 strSql = strSql & " FROM db_event.dbo.tbl_partner_eventStateLog as l "
 strSql = strSql & "  left outer join db_user.dbo.tbl_user_c as c on l.regid = c.userid " 
 strSql = strSql & "  left outer join db_partner.dbo.tbl_user_tenbyten as t on l.regid = t.userid "
 strSql = strSql & " where evt_code = "&FevtCode&"  order by evtlog_code desc " 	  
 rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetEventLog = rsget.getrows()
end if
rsget.close
End Function

 public Function fnGetMakerid
 	dim strSql
 	strSql =" select brand from db_event.dbo.tbl_partner_event where evt_code ="&FevtCode
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetMakerid = rsget(0)
	end if
	rsget.close
End Function
End Class                                
%>

 