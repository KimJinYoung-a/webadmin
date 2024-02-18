<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품 엑셀 업로드 일괄 수정
' History : 2019.04.18 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemedit_temp_cls.asp"-->
<%
dim mode, i, j, sqlStr, chk_idx, chk_idx_fail, sqldb, tempitemarr, tempitemoptionarr, tempitemoptiondetailarr, adminid
dim tempidx, realitemid, tempmileage, tempitemid
	mode 			= requestCheckVar(request("mode"),32)
	chk_idx 		= request("chk_idx")
	chk_idx_fail 		= request("chk_idx_fail")
    adminid = session("ssBctId")

dim refip
    refip = request.ServerVariables("HTTP_REFERER")

' 상품수정 업로드성공 상품 삭제
if (mode = "delitem") then
	sqlStr = " update db_temp.dbo.tbl_item_edit_temp" & vbcrlf
	sqlStr = sqlStr & " set isusing = 'N' where"
	sqlStr = sqlStr & " ordertempstatus <> 9 and ordertempstatus =1 and idx in ("& chk_idx &") and isusing = 'Y' "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('정상적으로 삭제 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

' 상품수정 업로드실패 상품 삭제
elseif (mode = "delitem_fail") then
	sqlStr = " update db_temp.dbo.tbl_item_edit_temp" & vbcrlf
	sqlStr = sqlStr & " set isusing = 'N' where"
	sqlStr = sqlStr & " ordertempstatus <> 9 and ordertempstatus =0 and idx in ("& chk_idx_fail &") and isusing = 'Y' "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('삭제 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

' 상품수정 실제 적용
elseif mode="edittemporder" then
    sqldb = " from db_temp.dbo.tbl_item_edit_temp t with (nolock)" & vbcrlf
    sqldb = sqldb & " join db_item.dbo.tbl_item i with (nolock)" & vbcrlf
    sqldb = sqldb & " 	on t.itemid = i.itemid" & vbcrlf
    sqldb = sqldb & " 	and i.sailyn='N'" & vbcrlf      ' 할인중인 상품 제낌
    sqldb = sqldb & " join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
    sqldb = sqldb & " 	on i.makerid=c.userid" & vbcrlf
    sqldb = sqldb & " left Join [db_item].[dbo].tbl_item_contents Ct with (nolock)"
    sqldb = sqldb & "     on t.itemid=Ct.itemid"
    sqldb = sqldb & " where t.isusing='Y'" & vbcrlf
    sqldb = sqldb & " and t.ordertempstatus=1" & vbcrlf
	sqldb = sqldb & " and t.idx in ("& chk_idx &")" & vbcrlf

	sqlStr = " update i" & vbcrlf
    sqlStr = sqlStr & " set i.lastupdate = getdate()" & vbcrlf
    sqlStr = sqlStr & " , i.itemname=convert(varchar(64),t.itemname)" & vbcrlf
    sqlStr = sqlStr & " , i.orgprice=t.orgprice" & vbcrlf
    sqlStr = sqlStr & " , i.sellcash=t.orgprice" & vbcrlf
    sqlStr = sqlStr & " , i.frontmakerid=t.frontmakerid" & vbcrlf
    sqlStr = sqlStr & sqldb

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	sqlStr = " update ct" & vbcrlf
    sqlStr = sqlStr & " set Ct.isbn13=t.isbn13" & vbcrlf
    sqlStr = sqlStr & sqldb

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	sqlStr = " update i" & vbcrlf
    sqlStr = sqlStr & " set i.lastupdate = getdate()" & vbcrlf
    sqlStr = sqlStr & " , i.buycash=round( (isnull(i.sellcash,0)*(100-IsNULL(c.defaultmargine,100))/100) ,0)" & vbcrlf
    sqlStr = sqlStr & " , i.orgsuplycash=round( (isnull(i.orgprice,0)*(100-IsNULL(c.defaultmargine,100))/100) ,0)" & vbcrlf
    sqlStr = sqlStr & sqldb

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	sqlStr = " update t" & vbcrlf
    sqlStr = sqlStr & " set t.ordertempstatus = 9" & vbcrlf
    sqlStr = sqlStr & sqldb

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('적용 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

' 상품신규등록 업로드성공 상품 삭제
elseif (mode = "delregitem") then
	sqlStr = " update db_temp.dbo.tbl_item_reg_temp" & vbcrlf
	sqlStr = sqlStr & " set isusing = 'N' where"
	sqlStr = sqlStr & " ordertempstatus <> 9 and ordertempstatus =1 and idx in ("& chk_idx &") and isusing = 'Y' "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('정상적으로 삭제 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

' 상품신규등록 업로드실패 상품 삭제
elseif (mode = "delregitem_fail") then
	sqlStr = " update db_temp.dbo.tbl_item_reg_temp" & vbcrlf
	sqlStr = sqlStr & " set isusing = 'N' where"
	sqlStr = sqlStr & " ordertempstatus <> 9 and ordertempstatus =0 and idx in ("& chk_idx_fail &") and isusing = 'Y' "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('삭제 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

' 상품신규등록 실제 적용
elseif mode="regtemporder" then
    sqlStr = "create table #tempitem(" & vbcrlf
    sqlStr = sqlStr & " idx int NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,tempitemid int NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,tempitemoption nvarchar(4) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,makerid nvarchar(32) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,dispcatecode bigint NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,itemname nvarchar(64) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,orgprice money NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,buycash money NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,mwdiv nvarchar(1) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,deliverytype nvarchar(1) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,itemoptionname nvarchar(96) NULL" & vbcrlf
    sqlStr = sqlStr & " ,barcode nvarchar(20) NULL" & vbcrlf
    sqlStr = sqlStr & " ,upchemanagecode nvarchar(32) NULL" & vbcrlf
    sqlStr = sqlStr & " ,realitemid int NULL" & vbcrlf
    sqlStr = sqlStr & " ,realitemoption nvarchar(4) NULL" & vbcrlf
    sqlStr = sqlStr & " ,brandname nvarchar(128) NULL" & vbcrlf
    sqlStr = sqlStr & " ,itemrackcode nvarchar(4) NULL" & vbcrlf
    sqlStr = sqlStr & " ,cate_large nvarchar(3) NULL" & vbcrlf
    sqlStr = sqlStr & " ,cate_mid nvarchar(3) NULL" & vbcrlf
    sqlStr = sqlStr & " ,cate_small nvarchar(3) NULL" & vbcrlf
    sqlStr = sqlStr & " ,buyitemname nvarchar(64) NULL" & vbcrlf
    sqlStr = sqlStr & " ,buyitemoptionname nvarchar(96) NULL" & vbcrlf
    sqlStr = sqlStr & " ,buycurrencyUnit nvarchar(16) NULL" & vbcrlf
    sqlStr = sqlStr & " ,buyitemprice money NULL" & vbcrlf
    sqlStr = sqlStr & " ,sourcearea nvarchar(128) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,makername nvarchar(64) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,keywords nvarchar(512) NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,volX float NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,volY float NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,volZ float NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,itemWeight int NOT NULL" & vbcrlf
    sqlStr = sqlStr & " ,frontmakerid nvarchar(32) NULL" & vbcrlf
    sqlStr = sqlStr & " )" & vbcrlf
    sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_idx ON #tempitem(idx ASC)" & vbcrlf
    sqlStr = sqlStr & " insert into #tempitem(" & vbcrlf
    sqlStr = sqlStr & " idx,tempitemid,tempitemoption,makerid,dispcatecode,itemname,orgprice,buycash" & vbcrlf
    sqlStr = sqlStr & " ,mwdiv,deliverytype,itemoptionname,barcode,upchemanagecode,realitemid,realitemoption" & vbcrlf
    sqlStr = sqlStr & " ,cate_large,cate_mid,cate_small,buyitemname,buyitemoptionname,buycurrencyUnit,buyitemprice" & vbcrlf
    sqlStr = sqlStr & " ,sourcearea,makername,keywords,volX,volY,volZ,itemWeight,frontmakerid)" & vbcrlf
    sqlStr = sqlStr & "     select top 100" & vbcrlf    ' 100개씩 제한
    sqlStr = sqlStr & "     idx,tempitemid,tempitemoption,makerid,dispcatecode,itemname,orgprice,buycash" & vbcrlf
    sqlStr = sqlStr & "     ,mwdiv,deliverytype,itemoptionname,barcode,upchemanagecode,realitemid,realitemoption" & vbcrlf
    sqlStr = sqlStr & "     ,cate_large,cate_mid,cate_small,buyitemname,buyitemoptionname,buycurrencyUnit,buyitemprice" & vbcrlf
    sqlStr = sqlStr & "     ,sourcearea,makername,keywords,volX,volY,volZ,itemWeight,t.frontmakerid" & vbcrlf
    sqlStr = sqlStr & "     from db_temp.dbo.tbl_item_reg_temp t with (readuncommitted)" & vbcrlf
    sqlStr = sqlStr & "     where t.isusing='Y'" & vbcrlf
    sqlStr = sqlStr & "     and t.ordertempstatus=1" & vbcrlf
    sqlStr = sqlStr & "     and t.regadminid='"& adminid &"'" & vbcrlf
	sqlStr = sqlStr & "     and t.idx in ("& chk_idx &")" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

    ' 브랜드 이름 넣기, 상품랙코드 넣기
    sqlStr = "update t set t.brandname=c.socname, t.itemrackcode=(case when isnull(c.prtidx,'9999')='' then '9999' else isnull(c.prtidx,'9999') end)" & vbcrlf
    sqlStr = sqlStr & " from #tempitem t" & vbcrlf
    sqlStr = sqlStr & " join [db_user].[dbo].tbl_user_c c" & vbcrlf
    sqlStr = sqlStr & "     on t.makerid=c.userid" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

    sqlStr = "SELECT" & vbcrlf
    sqlStr = sqlStr & " idx,tempitemid,tempitemoption,makerid,dispcatecode,itemname,orgprice,buycash" & vbcrlf
    sqlStr = sqlStr & " ,mwdiv,deliverytype,itemoptionname,barcode,upchemanagecode,realitemid,realitemoption" & vbcrlf
    sqlStr = sqlStr & " FROM #tempitem" & vbcrlf
    sqlStr = sqlStr & " where tempitemoption='0000'" & vbcrlf       ' 옵션없음

    'response.write sqlStr & "<Br>"
    rsget.open sqlStr,dbget
    If Not rsget.Eof Then
        tempitemarr = rsget.getrows()
    End If
    rsget.close

    ' 옵션없음 처리
    if isarray(tempitemarr) then
        for i = 0 to ubound(tempitemarr,2)
        tempidx = tempitemarr(0,i)
        tempitemid = tempitemarr(1,i)
        tempmileage = CLng(fix(tempitemarr(6,i)*0.005))

        dbget.beginTrans

        '// 제품번호를 받는다 //
		sqlStr = "Select max(itemid) as maxitemid from [db_item].[dbo].tbl_item" & vbcrlf

        'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			realitemid = rsget("maxitemid") + 1
		end if
		rsget.close
		
        '// 관리카테고리 저장 : 등록시 기본 1 CateGory만 **//
		sqlStr = "Insert into [db_item].dbo.tbl_Item_category(" & vbcrlf
		sqlStr = sqlStr & " itemid,code_large,code_mid,code_small,code_div)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.cate_large, t.cate_mid, t.cate_small, 'D'" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        '// 상품 데이터 입력 //
        sqlStr = "insert into db_item.dbo.tbl_item (" & vbcrlf
        sqlStr = sqlStr & " itemid,cate_large,cate_mid,cate_small,itemdiv, makerid,itemname" & vbcrlf
        sqlStr = sqlStr & " , sellcash ,buycash, orgprice, orgsuplycash, mileage, sellyn, deliverytype" & vbcrlf
        sqlStr = sqlStr & " , limityn,limitno,limitsold,limitdispyn,orderMinNum,orderMaxNum" & vbcrlf
        sqlStr = sqlStr & " , vatinclude, pojangok, deliverarea, deliverfixday, mwdiv, itemscore, upchemanagecode" & vbcrlf
        sqlStr = sqlStr & " , itemrackcode, tenOnlyYn, adultType, deliverOverseas , sellSTDate, optioncnt, dispcate1, brandname, itemWeight" & vbcrlf
        sqlStr = sqlStr & " , frontmakerid)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.cate_large, t.cate_mid, t.cate_small,'01', t.makerid, convert(varchar(64),t.itemname)" & vbcrlf
        sqlStr = sqlStr & "     , t.orgprice, t.buycash, t.orgprice, t.buycash, "& tempmileage &", 'N', t.deliverytype" & vbcrlf
        sqlStr = sqlStr & "     ,'N',0,0,'N',1,100" & vbcrlf
        sqlStr = sqlStr & "     ,'Y','N','','', t.mwdiv, 0, convert(varchar(32),t.upchemanagecode)" & vbcrlf
        sqlStr = sqlStr & "     ,t.itemrackcode,'N',0,'N',NULL,0, left(t.dispcatecode,3), t.brandname, t.itemWeight" & vbcrlf
        sqlStr = sqlStr & "     ,t.frontmakerid" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ''-- 상품 컨텐츠
        sqlStr = "insert into [db_item].[dbo].tbl_item_Contents (" & vbCrlf
        sqlStr = sqlStr & " itemid, keywords, sourcearea, makername,itemsource,itemsize,usinghtml,itemcontent" & vbCrlf
        sqlStr = sqlStr & " , ordercomment,designercomment, requireMakeDay,infoDiv,safetyYn,isbn13,isbn10,isbn_sub" & vbCrlf
        sqlStr = sqlStr & " , freight_min,freight_max, sourcekind)"  & vbCrlf	',safetyDiv,safetyNum
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.keywords,convert(varchar(128),t.sourcearea),t.makername,NULL,NULL,'N',NULL" & vbcrlf
        sqlStr = sqlStr & "     ,NULL,NULL,0,NULL,'N',NULL,NULL,NULL" & vbcrlf
        sqlStr = sqlStr & "     ,NULL,NULL,NULL" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		'// 전시카테고리 넣기 //
        sqlStr = "insert into db_item.dbo.tbl_display_cate_item (" & vbcrlf
        sqlStr = sqlStr & " catecode, itemid, depth, sortNo, isDefault)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     t.dispcatecode, "& realitemid &", convert(int,len(t.dispcatecode)/3), 9999, 'y'" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and isnull(t.dispcatecode,'')<>''" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ' 업체관리코드,범용바코드 등록
        sqlStr = "insert into db_item.dbo.tbl_item_option_stock (" & vbcrlf
        sqlStr = sqlStr & " itemgubun,itemid,itemoption,barcode,limitsellyn,limitsellno,limitsoldno,currstockno, upchemanagecode)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     '10', "& realitemid &",'0000',t.barcode,'N',0,0,0,convert(varchar(32),t.upchemanagecode)" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.barcode,'')<>'' or isnull(t.upchemanagecode,'')<>'')" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        sqlStr = "update t set t.realitemid="& realitemid &", t.realitemoption='0000'" & vbcrlf
        sqlStr = sqlStr & " from #tempitem t" & vbcrlf
        sqlStr = sqlStr & " where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' 업체매입용 정보 입력
        sqlStr = "insert into db_shop.dbo.tbl_buy_item(" & vbcrlf
        sqlStr = sqlStr & " itemgubun, buyitemid, itemoption, makerid, buyitemname, buyitemoptionname, buyitemprice, currencyUnit, isusing, regdate, updt)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     '10',t.realitemid,t.realitemoption,t.makerid,convert(varchar(64),t.buyitemname),convert(varchar(96),t.buyitemoptionname),t.buyitemprice,t.buycurrencyUnit,'Y',getdate(),getdate()" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        'sqlStr = sqlStr & "     left join db_shop.dbo.tbl_buy_item bi with (readuncommitted)" & vbcrlf
        'sqlStr = sqlStr & "         on bi.itemgubun='10'" & vbcrlf
        'sqlStr = sqlStr & "         and t.realitemid = bi.buyitemid" & vbcrlf
        'sqlStr = sqlStr & "         and t.realitemoption = bi.itemoption" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.buyitemname,'')<>'' or isnull(t.buycurrencyUnit,'')<>'' or isnull(t.buyitemprice,0)<>0)" & vbcrlf
        'sqlStr = sqlStr & "     and bi.itemgubun is null" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' 무게와 사이즈 입력
		sqlStr = "INSERT INTO db_item.dbo.tbl_item_Volumn (" & VbCrlf
        sqlStr = sqlStr & " itemid, itemoption, itemWeight, volX, volY, volZ, regdate, lastupdate)" & VbCrlf
		sqlStr = sqlStr & "     select" & VbCrlf
        sqlStr = sqlStr & "     t.realitemid,t.realitemoption,t.itemWeight,t.volX,t.volY,t.volZ,getdate(),getdate()" & VbCrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.itemWeight,0)<>0 or isnull(t.volX,0)<>0 or isnull(t.volY,0)<>0or isnull(t.volZ,0)<>0)" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' (구)무게와 사이즈 입력
		sqlStr = "INSERT INTO db_item.dbo.tbl_item_pack_Volumn (" & VbCrlf
        sqlStr = sqlStr & " itemid, volX, volY, volZ, regdate, lastupdt)" & VbCrlf
		sqlStr = sqlStr & "     select" & VbCrlf
        sqlStr = sqlStr & "     t.realitemid,t.volX,t.volY,t.volZ,getdate(),getdate()" & VbCrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.volX,0)<>0 or isnull(t.volY,0)<>0or isnull(t.volZ,0)<>0)" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        If Err.Number = 0 Then
            dbget.CommitTrans
            'dbget.RollBackTrans

            ' 상품임시등록 테이블에 실상품번호와 옵션번호 엎어침
            sqlStr = "update it set it.realitemid=t.realitemid, it.realitemoption=t.realitemoption, it.ordertempstatus = 9" & vbcrlf
            sqlStr = sqlStr & " from #tempitem t" & vbcrlf
            sqlStr = sqlStr & " join db_temp.dbo.tbl_item_reg_temp it" & vbcrlf
            sqlStr = sqlStr & "     on t.idx=it.idx" & vbcrlf
            sqlStr = sqlStr & "     and t.realitemid is not null and t.realitemoption is not null" & vbcrlf
            sqlStr = sqlStr & " where t.idx="& tempidx &"" & vbcrlf

            'response.write sqlStr & "<br>"
            dbget.execute sqlStr
        Else
            dbget.RollBackTrans
            response.write "<script type='text/javascript'>"
            response.write "    location.href ='about:blank';"
            response.write "    alert('임시상품번호("& tempitemid &") 처리중 에러가 발생했습니다.');"
            response.write "</script>"
            dbget.close() : response.end
        End If
        next
    end if

    sqlStr = "select *" & vbcrlf
    sqlStr = sqlStr & " from (" & vbcrlf
    sqlStr = sqlStr & " 	select idx,tempitemid,makerid,dispcatecode,itemname,orgprice,buycash,mwdiv,deliverytype,cate_large,cate_mid,cate_small" & vbcrlf
    sqlStr = sqlStr & " 	, RANK() Over (partition by tempitemid order by tempitemoption asc) as LastRank" & vbcrlf
    sqlStr = sqlStr & " 	from #tempitem" & vbcrlf
    sqlStr = sqlStr & "     where tempitemoption<>'0000'" & vbcrlf       ' 옵션있음
    sqlStr = sqlStr & " ) as t" & vbcrlf
    sqlStr = sqlStr & " where LastRank=1" & vbcrlf

    'response.write sqlStr & "<Br>"
    rsget.open sqlStr,dbget
    If Not rsget.Eof Then
        tempitemoptionarr = rsget.getrows()
    End If
    rsget.close

    ' 옵션있음 처리
    i = 0
    if isarray(tempitemoptionarr) then
        for i = 0 to ubound(tempitemoptionarr,2)
        tempidx = tempitemoptionarr(0,i)
        tempitemid = tempitemoptionarr(1,i)
        tempmileage = CLng(fix(tempitemoptionarr(6,i)*0.005))

        dbget.beginTrans

        '// 제품번호를 받는다 //
		sqlStr = "Select max(itemid) as maxitemid from [db_item].[dbo].tbl_item" & vbcrlf

        'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			realitemid = rsget("maxitemid") + 1
		end if
		rsget.close
		
        '// 관리카테고리 저장 : 등록시 기본 1 CateGory만 **//
		sqlStr = "Insert into [db_item].dbo.tbl_Item_category(" & vbcrlf
		sqlStr = sqlStr & " itemid,code_large,code_mid,code_small,code_div)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.cate_large, t.cate_mid, t.cate_small, 'D'" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        '// 상품 데이터 입력 //
        sqlStr = "insert into db_item.dbo.tbl_item (" & vbcrlf
        sqlStr = sqlStr & " itemid,cate_large,cate_mid,cate_small,itemdiv, makerid,itemname" & vbcrlf
        sqlStr = sqlStr & " , sellcash ,buycash, orgprice, orgsuplycash, mileage, sellyn, deliverytype" & vbcrlf
        sqlStr = sqlStr & " , limityn,limitno,limitsold,limitdispyn,orderMinNum,orderMaxNum" & vbcrlf
        sqlStr = sqlStr & " , vatinclude, pojangok, deliverarea, deliverfixday, mwdiv, itemscore" & vbcrlf
        sqlStr = sqlStr & " , itemrackcode, tenOnlyYn, adultType, deliverOverseas , sellSTDate, dispcate1, brandname, itemWeight" & vbcrlf
        sqlStr = sqlStr & " , frontmakerid)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.cate_large, t.cate_mid, t.cate_small,'01', t.makerid, convert(varchar(64),t.itemname)" & vbcrlf
        sqlStr = sqlStr & "     , t.orgprice, t.buycash, t.orgprice, t.buycash, "& tempmileage &", 'N', t.deliverytype" & vbcrlf
        sqlStr = sqlStr & "     ,'N',0,0,'N',1,100" & vbcrlf
        sqlStr = sqlStr & "     ,'Y','N','','', t.mwdiv, 0" & vbcrlf
        sqlStr = sqlStr & "     ,t.itemrackcode,'N',0,'N',NULL, left(t.dispcatecode,3), t.brandname, t.itemWeight" & vbcrlf
        sqlStr = sqlStr & "     ,t.frontmakerid" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ''-- 상품 컨텐츠
        sqlStr = "insert into [db_item].[dbo].tbl_item_Contents (" & vbCrlf
        sqlStr = sqlStr & " itemid, keywords, sourcearea, makername,itemsource,itemsize,usinghtml,itemcontent" & vbCrlf
        sqlStr = sqlStr & " , ordercomment,designercomment, requireMakeDay,infoDiv,safetyYn,isbn13,isbn10,isbn_sub" & vbCrlf
        sqlStr = sqlStr & " , freight_min,freight_max, sourcekind)"  & vbCrlf	',safetyDiv,safetyNum
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &", t.keywords,convert(varchar(128),t.sourcearea),t.makername,NULL,NULL,'N',NULL" & vbcrlf
        sqlStr = sqlStr & "     ,NULL,NULL,0,NULL,'N',NULL,NULL,NULL" & vbcrlf
        sqlStr = sqlStr & "     ,NULL,NULL,NULL" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		'// 전시카테고리 넣기 //
        sqlStr = "insert into db_item.dbo.tbl_display_cate_item (" & vbcrlf
        sqlStr = sqlStr & " catecode, itemid, depth, sortNo, isDefault)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     t.dispcatecode, "& realitemid &", convert(int,len(t.dispcatecode)/3), 9999, 'y'" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and isnull(t.dispcatecode,'')<>''" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ' 단일옵션등록
        sqlStr = "insert into db_item.dbo.tbl_item_option (" & vbcrlf
        sqlStr = sqlStr & " itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold)" & vbcrlf        
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     "& realitemid &",t.tempitemoption,'',convert(varchar(96),t.itemoptionname),'Y','N','N',0,0" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.tempitemid="& tempitemid &"" & vbcrlf
        sqlStr = sqlStr & "     and isnull(t.itemoptionname,'')<>'' and isnull(t.tempitemoption,'')<>''" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ''옵션 총수 저장
        sqlStr = " update i"
        sqlStr = sqlStr & " set i.optioncnt=IsNULL(T.cnt,0)"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
        sqlStr = sqlStr & " join ("
        sqlStr = sqlStr & "     select itemid, count(itemid) as cnt"
        sqlStr = sqlStr & "     from db_item.dbo.tbl_item_option"
        sqlStr = sqlStr & "     where itemid = "& realitemid &" and isusing = 'Y'"
        sqlStr = sqlStr & "     group by itemid"
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & "     on i.itemid=t.itemid"
        sqlStr = sqlStr & " where i.itemid = "& realitemid &""

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        ' 업체관리코드,범용바코드 등록
        sqlStr = "insert into db_item.dbo.tbl_item_option_stock (" & vbcrlf
        sqlStr = sqlStr & " itemgubun,itemid,itemoption,barcode,limitsellyn,limitsellno,limitsoldno,currstockno, upchemanagecode)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     '10', "& realitemid &",t.tempitemoption,t.barcode,'N',0,0,0,convert(varchar(32),t.upchemanagecode)" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.tempitemid="& tempitemid &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.barcode,'')<>'' or isnull(t.upchemanagecode,'')<>'')" & vbcrlf

        'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

        sqlStr = "update t set t.realitemid="& realitemid &", t.realitemoption=t.tempitemoption" & vbcrlf
        sqlStr = sqlStr & " from #tempitem t" & vbcrlf
        sqlStr = sqlStr & " where t.tempitemid="& tempitemid &"" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' 업체매입용 정보 입력
        sqlStr = "insert into db_shop.dbo.tbl_buy_item(" & vbcrlf
        sqlStr = sqlStr & " itemgubun, buyitemid, itemoption, makerid, buyitemname, buyitemoptionname, buyitemprice, currencyUnit, isusing, regdate, updt)" & vbcrlf
        sqlStr = sqlStr & "     select" & vbcrlf
        sqlStr = sqlStr & "     '10',t.realitemid,t.realitemoption,t.makerid,convert(varchar(64),t.buyitemname),convert(varchar(96),t.buyitemoptionname),t.buyitemprice,t.buycurrencyUnit,'Y',getdate(),getdate()" & vbcrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        'sqlStr = sqlStr & "     left join db_shop.dbo.tbl_buy_item bi with (readuncommitted)" & vbcrlf
        'sqlStr = sqlStr & "         on bi.itemgubun='10'" & vbcrlf
        'sqlStr = sqlStr & "         and t.realitemid = bi.buyitemid" & vbcrlf
        'sqlStr = sqlStr & "         and t.realitemoption = bi.itemoption" & vbcrlf
        sqlStr = sqlStr & "     where t.tempitemid="& tempitemid &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.buyitemname,'')<>'' or isnull(t.buycurrencyUnit,'')<>'' or isnull(t.buyitemprice,0)<>0)" & vbcrlf
        'sqlStr = sqlStr & "     and bi.itemgubun is null" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' 무게와 사이즈 입력
		sqlStr = "INSERT INTO db_item.dbo.tbl_item_Volumn (" & VbCrlf
        sqlStr = sqlStr & " itemid, itemoption, itemWeight, volX, volY, volZ, regdate, lastupdate)" & VbCrlf
		sqlStr = sqlStr & "     select" & VbCrlf
        sqlStr = sqlStr & "     t.realitemid,t.realitemoption,t.itemWeight,t.volX,t.volY,t.volZ,getdate(),getdate()" & VbCrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.tempitemid="& tempitemid &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.itemWeight,0)<>0 or isnull(t.volX,0)<>0 or isnull(t.volY,0)<>0or isnull(t.volZ,0)<>0)" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        ' (구)무게와 사이즈 입력
		sqlStr = "INSERT INTO db_item.dbo.tbl_item_pack_Volumn (" & VbCrlf
        sqlStr = sqlStr & " itemid, volX, volY, volZ, regdate, lastupdt)" & VbCrlf
		sqlStr = sqlStr & "     select" & VbCrlf
        sqlStr = sqlStr & "     t.realitemid,t.volX,t.volY,t.volZ,getdate(),getdate()" & VbCrlf
        sqlStr = sqlStr & "     FROM #tempitem t" & vbcrlf
        sqlStr = sqlStr & "     where t.idx="& tempidx &"" & vbcrlf
        sqlStr = sqlStr & "     and (isnull(t.volX,0)<>0 or isnull(t.volY,0)<>0or isnull(t.volZ,0)<>0)" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr

        If Err.Number = 0 Then
            dbget.CommitTrans
            'dbget.RollBackTrans

            ' 상품임시등록 테이블에 실상품번호와 옵션번호 엎어침
            sqlStr = "update it set it.realitemid=t.realitemid, it.realitemoption=t.realitemoption, it.ordertempstatus = 9" & vbcrlf
            sqlStr = sqlStr & " from #tempitem t" & vbcrlf
            sqlStr = sqlStr & " join db_temp.dbo.tbl_item_reg_temp it" & vbcrlf
            sqlStr = sqlStr & "     on t.idx=it.idx" & vbcrlf
            sqlStr = sqlStr & "     and t.realitemid is not null and t.realitemoption is not null" & vbcrlf
            sqlStr = sqlStr & " where t.tempitemid="& tempitemid &"" & vbcrlf

            'response.write sqlStr & "<br>"
            dbget.execute sqlStr
        Else
            dbget.RollBackTrans
            response.write "<script type='text/javascript'>"
            response.write "    location.href ='about:blank';"
            response.write "    alert('임시상품번호("& tempitemid &") 처리중 에러가 발생했습니다.');"
            response.write "</script>"
            dbget.close() : response.end
        End If
        next
    end if

    sqlStr = sqlStr & " drop table #tempitem" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
    response.write "    location.href ='about:blank';"
    response.write "    alert('적용 되었습니다.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end
else
	response.write "<script type='text/javascript'>"
    response.write "    alert('구분자가 없습니다.');"
    response.write "</script>"
    dbget.close() : response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->