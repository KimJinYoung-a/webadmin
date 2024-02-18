<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 180
%>
<%
'###########################################################
' Description :  오프라인매장 정리대상 상품 포함 브랜드
' History : 2011.08 서동석 생성
'			2017.04.16 한용민 수정(보안관련처리)
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%
Dim params : params = request.Form("params")
Dim shopid : shopid = requestCheckVar(request.Form("shopid"),32)
Dim makerid : makerid = requestCheckVar(request.Form("makerid"),32)
Dim lossDate : lossDate = requestCheckVar(request.Form("lossDate"),30)
Dim losstype : losstype = requestCheckVar(request.Form("losstype"),10)

Dim cksel : cksel = request.Form("cksel") + ","
Dim AssignrealcheckErr : AssignrealcheckErr = request.Form("AssignrealcheckErr") + ","
Dim shopitemprice : shopitemprice = request.Form("shopitemprice") + ","
Dim shopbuyprice : shopbuyprice = request.Form("shopbuyprice") + ","
Dim shopsuplycash : shopsuplycash = request.Form("shopsuplycash") + ","
Dim itemgubun : itemgubun = request.Form("itemgubun") + ","
Dim itemid    : itemid = request.Form("itemid") + ","
Dim itemoption  : itemoption = request.Form("itemoption") + ","
Dim cType : cType = requestCheckVar(request.Form("cType"),10)
Dim mode : mode = requestCheckVar(request.Form("mode"),32)

Dim iLOSSBRandID , iLOSSCAUSE
'IF (cType="C") then
'    iLOSSBRandID = "shopstockmodify"
'    iLOSSCAUSE   = "재고조정"
'ELSEIF (cType="L") then
'    iLOSSBRandID = "shopitemloss"
'    iLOSSCAUSE   = "로스처리"
'ELSE
'    iLOSSBRandID = "shopstockmodify"
'    iLOSSCAUSE   = "재고조정"
'END IF

iLOSSBRandID = "shopitemloss"
IF (losstype="M") then
    iLOSSCAUSE   = "로스처리(정산미반영)"
ELSEIF (losstype="S") then
    iLOSSCAUSE   = "샘플폐기(정산미반영)"
	iLOSSBRandID = "shopitemsample"
ELSEIF (losstype="L") then
    iLOSSCAUSE   = "로스처리(정산반영)"
END IF


Dim sqlStr, idx, i, cnt, vix

rw shopid
rw makerid
rw cksel
rw AssignrealcheckErr
rw shopsuplycash


cksel     = split(cksel,",")
itemgubun = split(itemgubun,",")
itemid    = split(itemid,",")
itemoption= split(itemoption,",")
AssignrealcheckErr = split(AssignrealcheckErr,",")
shopsuplycash      = split(shopsuplycash,",")
shopbuyprice        = split(shopbuyprice,",")
shopitemprice       = split(shopitemprice,",")

if IsArray(cksel) then
    cnt = Ubound(cksel)
else
    cnt = 0
end if

Dim retURL : retURL="/admin/offshop/stock/OutItemListByBrand.asp?"&params

if (mode="stockupArr") then
    for i=0 to cnt
	    vix = cksel(i)
	    if (vix<>"") then
    	    If (itemgubun(vix)<>"") and (itemid(vix)<>"") and (itemoption(vix)<>"")  then
        		''-1 월말 업데이트
                sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & requestCheckVar(Trim(itemgubun(vix)),2) & "'," & requestCheckVar(Trim(itemid(vix)),10) & ",'" & requestCheckVar(Trim(itemoption(vix)),4) & "'"
                dbget.Execute sqlStr

                ''-1 일별 업데이트
                sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & requestCheckVar(Trim(itemgubun(vix)),2) & "'," & requestCheckVar(Trim(itemid(vix)),10) & ",'" & requestCheckVar(Trim(itemoption(vix)),4) & "'"
                dbget.Execute sqlStr

                ''rw  sqlStr
            end if
        end if

	next

	response.write "<script type='text/javascript'>alert('처리 되었습니다.');</script>"
    response.write "<script type='text/javascript'>location.href='"&retURL&"';</script>"
    response.end

end if

''2개월 이전 자료는 입력 못함..
Dim STOCKBASEDATE : STOCKBASEDATE = Left(dateAdd("m",-1,now()),7) + "-01"
Dim isPreMonth : isPreMonth =FALSE

IF (CDate(lossDate)<CDate(STOCKBASEDATE)) THEN
   isPreMonth = TRUE
   if (Not C_ADMIN_AUTH) and (Not C_OFF_AUTH) then
       response.write STOCKBASEDATE & " 이전 날짜로 설정 불가 - 관리자 메뉴"
       dbget.Close() : response.end
    end if
End if

''' 로스 출고 입력
''isreq 입고요청. Flag , isbaljuExists 'Y'
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists,comment)"
	sqlStr = sqlStr + " values('" + iLOSSBRandID + "'"
	sqlStr = sqlStr + " ,'" + shopid + "'"
	sqlStr = sqlStr + " ,'999'"
	sqlStr = sqlStr + " ,'008'"
	sqlStr = sqlStr + " ,'" + lossDate + "'"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " ,'99'"
	sqlStr = sqlStr + " ,'"&iLOSSCAUSE&"'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,'N'"
	sqlStr = sqlStr + " ,'실사오차 "&iLOSSCAUSE&"'"
	sqlStr = sqlStr + " )"

	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close


	for i=0 to cnt
	    vix = cksel(i)
	    if (vix<>"") then
    	    If (itemgubun(vix)<>"") and (itemid(vix)<>"") and (itemoption(vix)<>"") and (shopsuplycash(vix)<>"") and (shopsuplycash(vix)<>"") then
        		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
        		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
        		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf  ''suplycash 매입가 , shopbuyprice 매장공급가
        		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
        		sqlStr = sqlStr + "'" + requestCheckVar(Trim(itemgubun(vix)),2) + "'," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(itemid(vix)),10) + "," + vbCrlf
        		sqlStr = sqlStr + "'" + requestCheckVar(Trim(itemoption(vix)),4) + "'," + vbCrlf
        		sqlStr = sqlStr + "'" + makerid + "'," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopitemprice(vix)),20) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopbuyprice(vix)),20) + "," + vbCrlf      '' 원매입가 매입가.
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopsuplycash(vix)),20) + "," + vbCrlf       '' 정산반영액
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(AssignrealcheckErr(vix)*-1),20) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(AssignrealcheckErr(vix)*-1),20) + vbCrlf
        		sqlStr = sqlStr + "" + ")"

        		dbget.Execute(sqlStr)


				if (losstype="S") Then
					'''샘플
					''sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekSample_Input '" & shopid & "','" & Trim(itemgubun(vix)) & "'," & Trim(itemid(vix)) & ",'" & Trim(itemoption(vix)) & "'," & AssignrealcheckErr(vix)*-1 & ",'" & session("ssBctID") & "','" & lossDate & "'"
					sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekSample_Input '" & shopid & "','" & requestCheckVar(Trim(itemgubun(vix)),2) & "'," & requestCheckVar(Trim(itemid(vix)),10) & ",'" & requestCheckVar(Trim(itemoption(vix)),4) & "'," & 0 & ",'" & session("ssBctID") & "','" & lossDate & "'"
				Else
					'''오차 차감
					sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekErr_Input '" & shopid & "','" & requestCheckVar(Trim(itemgubun(vix)),2) & "'," & requestCheckVar(Trim(itemid(vix)),10) & ",'" & requestCheckVar(Trim(itemoption(vix)),4) & "'," & requestCheckVar(AssignrealcheckErr(vix),20) & ",'" & lossDate & "','" & session("ssBctID") & "'"
				End If

                rw sqlStr
                dbget.Execute sqlStr

            end if
        end if

	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,statecd='8'"  + vbCrlf
	sqlStr = sqlStr + " ,execdt='"&lossDate&"'"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmdate=getdate()"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " ,comm_cd='"&losstype&"'"                                            '''로스타입
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " ,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " ,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	dbget.Execute(sqlStr)

    ''월별 재고 업데이트, 월말재고 업데이트.
    sqlStr = " Insert into db_summary.dbo.tbl_monthly_shopstock_summary"&VbCRLF
    sqlStr = sqlStr&" (shopid,itemgubun,itemid,itemoption,yyyymm,"&VbCRLF
    sqlStr = sqlStr&" logicsipgono,logicsreipgono,brandipgono,brandreipgono,sellno,resellno,"&VbCRLF
    sqlStr = sqlStr&" errsampleitemno,errbaditemno,errrealcheckno,sysstockno,realstockno,regdate,lastupdate)"&VbCRLF
    sqlStr = sqlStr&" select "&VbCRLF
    sqlStr = sqlStr&" m.shopid,d.itemgubun,d.shopitemid,d.itemoption,convert(varchar(7),m.execdt,21)"&VbCRLF
    sqlStr = sqlStr&" ,0,0,0,0,0,0,0,0,0,0,0,getdate(),getdate()"&VbCRLF
    sqlStr = sqlStr&" from [db_shop].[dbo].tbl_shop_ipchul_detail d"&VbCRLF
    sqlStr = sqlStr&" 	Join [db_shop].[dbo].tbl_shop_ipchul_master m"&VbCRLF
    sqlStr = sqlStr&" 	on m.idx=d.masteridx"&VbCRLF
    sqlStr = sqlStr&" 	left Join db_summary.dbo.tbl_monthly_shopstock_summary s"&VbCRLF
    sqlStr = sqlStr&" 	on s.yyyymm=convert(varchar(7),m.execdt,21)"&VbCRLF
    sqlStr = sqlStr&" 	and s.shopid=m.shopid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemgubun=d.itemgubun"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemid=d.shopitemid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemoption=d.itemoption"&VbCRLF
    sqlStr = sqlStr&" where d.masteridx="+ CStr(idx) &VbCRLF
    sqlStr = sqlStr&" and d.deleteyn='N'"&VbCRLF
    sqlStr = sqlStr&" and s.itemgubun is NULL"&VbCRLF
    dbget.Execute(sqlStr)

    sqlStr = " update S"&VbCRLF
    sqlStr = sqlStr&" set brandipgono=(Case when d.itemno>0 then d.itemno else s.brandipgono end)"&VbCRLF
    sqlStr = sqlStr&" ,brandreipgono=(Case when d.itemno<0 then d.itemno else s.brandreipgono end)"&VbCRLF
    sqlStr = sqlStr&" ,sysstockno=sysstockno+d.itemno"&VbCRLF

	if (losstype="S") Then
		sqlStr = sqlStr&" ,errsampleitemno=errsampleitemno-d.itemno"&VbCRLF
	Else
		sqlStr = sqlStr&" ,errrealcheckno=errrealcheckno-d.itemno"&VbCRLF
	End If

    sqlStr = sqlStr&" ,lastupdate=getdate()"&VbCRLF
    sqlStr = sqlStr&" from [db_shop].[dbo].tbl_shop_ipchul_detail d"&VbCRLF
    sqlStr = sqlStr&" 	Join [db_shop].[dbo].tbl_shop_ipchul_master m"&VbCRLF
    sqlStr = sqlStr&" 	on m.idx=d.masteridx"&VbCRLF
    sqlStr = sqlStr&" 	Join db_summary.dbo.tbl_monthly_shopstock_summary s"&VbCRLF
    sqlStr = sqlStr&" 	on s.yyyymm=convert(varchar(7),m.execdt,21)"&VbCRLF
    sqlStr = sqlStr&" 	and s.shopid=m.shopid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemgubun=d.itemgubun"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemid=d.shopitemid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemoption=d.itemoption"&VbCRLF
    sqlStr = sqlStr&" where d.masteridx="+ CStr(idx) &VbCRLF
    sqlStr = sqlStr&" and d.deleteyn='N'"&VbCRLF
	''rw sqlStr
    dbget.Execute(sqlStr)

    ''월말재고 업데이트 - 재고 있는내역만..
    sqlStr = " update S"&VbCRLF
    sqlStr = sqlStr&" set brandipgono=(Case when d.itemno>0 then d.itemno else s.brandipgono end)"&VbCRLF
    sqlStr = sqlStr&" ,brandreipgono=(Case when d.itemno<0 then d.itemno else s.brandreipgono end)"&VbCRLF
    sqlStr = sqlStr&" ,sysstockno=sysstockno+d.itemno"&VbCRLF

	if (losstype="S") Then
		sqlStr = sqlStr&" ,errsampleitemno=errsampleitemno-d.itemno"&VbCRLF
	Else
		sqlStr = sqlStr&" ,errrealcheckno=errrealcheckno-d.itemno"&VbCRLF
	End If

    ''sqlStr = sqlStr&" ,errrealcheckno=errrealcheckno-d.itemno"&VbCRLF
    sqlStr = sqlStr&" ,lastupdate=getdate()"&VbCRLF
    sqlStr = sqlStr&" from [db_shop].[dbo].tbl_shop_ipchul_detail d"&VbCRLF
    sqlStr = sqlStr&" 	Join [db_shop].[dbo].tbl_shop_ipchul_master m"&VbCRLF
    sqlStr = sqlStr&" 	on m.idx=d.masteridx"&VbCRLF
    sqlStr = sqlStr&" 	Join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s"&VbCRLF
    sqlStr = sqlStr&" 	on s.yyyymm>=convert(varchar(7),m.execdt,21)"&VbCRLF
    ''sqlStr = sqlStr&" 	on s.yyyymm='"&Left(lossDate,7)&"'"
    sqlStr = sqlStr&" 	and s.shopid=m.shopid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemgubun=d.itemgubun"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemid=d.shopitemid"&VbCRLF
    sqlStr = sqlStr&" 	and s.itemoption=d.itemoption"&VbCRLF
    sqlStr = sqlStr&" where d.masteridx="+ CStr(idx) &VbCRLF
    sqlStr = sqlStr&" and d.deleteyn='N'"&VbCRLF
	''rw sqlStr
	dbget.Execute(sqlStr)

    IF (isPreMonth) then
        for i=0 to cnt
    	    vix = cksel(i)
    	    if (vix<>"") then
        	    If (itemgubun(vix)<>"") and (itemid(vix)<>"") and (itemoption(vix)<>"")  then
            		''-1 월말 업데이트
                    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & Trim(itemgubun(vix)) & "'," & Trim(itemid(vix)) & ",'" & Trim(itemoption(vix)) & "'"
                    dbget.Execute sqlStr

                    ''-1 일별 업데이트
                    ''sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & Trim(itemgubun(vix)) & "'," & Trim(itemid(vix)) & ",'" & Trim(itemoption(vix)) & "'"
                    'dbget.Execute sqlStr

                    sqlStr = " update [db_summary].[dbo].tbl_current_shopstock_summary "
            		sqlStr = sqlStr&" set brandreipgono=brandreipgono + "&AssignrealcheckErr(vix)*-1
            		sqlStr = sqlStr&" ,sysstockno=sysstockno + "&AssignrealcheckErr(vix)*-1
            		sqlStr = sqlStr&" ,realstockno=realstockno + "&AssignrealcheckErr(vix)*-1
            		sqlStr = sqlStr&",lastupdate=getdate()"
            		sqlStr = sqlStr&"where shopid='"&shopid&"'"
            		sqlStr = sqlStr&"and itemgubun='"&Trim(itemgubun(vix))&"'"
            		sqlStr = sqlStr&"and itemid='"&Trim(itemid(vix))&"'"
            		sqlStr = sqlStr&"and itemoption='"&Trim(itemoption(vix))&"'"

		            dbget.Execute sqlStr
                end if
            end if

    	next
    ELSE
        ''재고 반영. //2개월 이내만 반영됨..
        sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",1"
        dbget.Execute(sqlStr)
	ENd IF

    response.write "<script type='text/javascript'>alert('처리 되었습니다.');</script>"
    response.write "<script type='text/javascript'>location.href='"&retURL&"';</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
