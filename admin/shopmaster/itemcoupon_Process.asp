<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
 Server.ScriptTimeout= 150	'5분

Dim otimer : otimer=Timer()
Dim pTimer : pTimer=otimer

function debugRwite (istep)
	if Application("Svr_info")="Dev" then
		rw istep&":"& FormatNumber(Timer()-pTimer,4)
		pTimer=Timer()
	end if
end function

function IsLastUpdateNotAssign(itemcouponidx)
    ''쿠폰에 속한 아이템 갯수가 2000개 이상일경우 lastupdate를 업데이트 안하기 위함.
    Dim sqlStr

    IsLastUpdateNotAssign = false

    sqlStr = "select count(*) as CNT from [db_item].[dbo].tbl_item_coupon_detail where itemcouponidx=" + CStr(itemcouponidx)
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    IsLastUpdateNotAssign =(rsget("CNT")>=2000)   ''상품수가 2000개 넘으면 오픈/종료시 lastupdate 처리 안함.
	end if
	rsget.Close

end function

function AplyItemCountUpdate(itemcouponidx)
	dim sqlStr
	''적용상품갯수 업데이트
	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set applyitemcount=IsNULL(T.cnt,0)" + VbCrlf
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select count(*) as cnt from [db_item].[dbo].tbl_item_coupon_detail where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
	
	''2차서버 2018/06/19 
	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	dblogicsget.Execute sqlStr
end function

function AplyToItem(itemcouponidx, chklastupdate)
	dim sqlStr
	dim ocouponGubun, oitemcoupontype, oitemcouponvalue, oitemcouponstartdate, oitemcouponexpiredate, openstate, currdatetime
	dim couponExpired
	dim resultCnt
    dim notUpdate

    '' 2010-10 추가
    ''상품 lastupdate 변경할지 여부
    ''chklastupdate 는 쿠폰 강제 오픈시/ 강제종료시만 체크함.
    ''notUpdate = true인경우 lastupdate 업데이트 하지 않음.
    notUpdate = false
    if (chklastupdate) then
        notUpdate = IsLastUpdateNotAssign(itemcouponidx)

        if (notUpdate) then response.write "변경상품수가 2000개이상이라 상품 lastupdate 변경안함.<br>"
    end if
	Call debugRwite("stepA-1")

	applyitemcount = 0
	couponExpired = false

	sqlStr = "select top 1 couponGubun, margintype, itemcoupontype, itemcouponvalue, openstate, applyitemcount,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
	sqlStr = sqlStr + " (case when (itemcouponstartdate>getdate()) or (itemcouponexpiredate<getdate()) then 'Y' else 'N' end ) as couponexpired, "
	sqlStr = sqlStr + " convert(varchar(19),getdate()) as currdatetime"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    ocouponGubun   = rsget("couponGubun")
		itemcoupontype = rsget("itemcoupontype")
		itemcouponvalue = rsget("itemcouponvalue")
		itemcouponstartdate = rsget("itemcouponstartdate")
		itemcouponexpiredate = rsget("itemcouponexpiredate")
		openstate = rsget("openstate")
		applyitemcount = rsget("applyitemcount")
		currdatetime = rsget("currdatetime")

		couponExpired = rsget("couponexpired")

		response.write "couponExpired :" + CStr(couponExpired) + "<br>"
	end if
	rsget.Close
	Call debugRwite("stepA-2")
    
	''발급대기중이거나 발급예약경우는 스킵.
	if ((openstate<>"7") and (openstate<>"9")) then exit function

	'' Naver 쿠폰인경우 lastupdate  2018/08/08
	if (ocouponGubun="V") and (Not notUpdate) then
		''배치로 처리하는것으로 변경하자.
		''EXEC db_AppWish.dbo.[sp_TEN_CP_tbl_item_coupon_master_detail_Change]
		' sqlStr = "EXEC db_AppWish.dbo.[sp_TEN_CP_tbl_item_coupon_master_detail_Change_By_CPnIDX] "&CStr(itemcouponidx)
		' dblogicsget.Execute sqlStr

		' ''2차서버 2018/06/19 
		' sqlStr = "update I "
		' sqlStr = sqlStr + " set lastupdate=getdate()"
		' sqlStr = sqlStr + " from [db_AppWish].[dbo].tbl_item I"
		' sqlStr = sqlStr + " 	Join [db_AppWish].[dbo].tbl_item_coupon_detail d"
		' sqlStr = sqlStr + " 	on I.itemid=d.itemid"
		' sqlStr = sqlStr + " where d.itemcouponidx=" + CStr(itemcouponidx)
		
		' dblogicsget.Execute sqlStr
	end if

	''타겟쿠폰, 지정쿠폰, 모바일전용인경우 스킵.
    if (ocouponGubun<>"C") then exit function

	''발급 종료된 쿠폰인경우 -> N로 변경
	if (openstate="9") or (couponExpired="Y") then

		sqlStr = "update [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " set itemcouponyn='N'"
		sqlStr = sqlStr + " ,itemcoupontype='1'"
		sqlStr = sqlStr + " ,itemcouponvalue=0"
		sqlStr = sqlStr + " ,curritemcouponidx=NULL"
		IF (Not notUpdate) then
		    sqlStr = sqlStr + " ,lastupdate=getdate()"
	    end if
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_detail"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.itemid=[db_item].[dbo].tbl_item_coupon_detail.itemid"

		'response.write sqlStr + "<br>"
		dbget.Execute sqlStr
		Call debugRwite("stepA-3")		

		''2차서버 2018/06/19 
	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	    dblogicsget.Execute sqlStr
		Call debugRwite("stepA-4")
	end if

	''상품이 삭제된경우 -> N로 변경 // 삭제한경우는 lastupdate쳐줌
	sqlStr = "update [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " set itemcouponyn='N'"
	sqlStr = sqlStr + " ,itemcoupontype='1'"
	sqlStr = sqlStr + " ,itemcouponvalue=0"
	sqlStr = sqlStr + " ,curritemcouponidx=NULL"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select i.itemid  "
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_coupon_detail d"
	sqlStr = sqlStr + " 	on d.itemcouponidx=" + CStr(itemcouponidx) + " and i.itemid=d.itemid "
	sqlStr = sqlStr + " 	where i.curritemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and d.itemcouponidx is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"

	'response.write sqlStr + "<br>"
		dbget.Execute sqlStr, resultCnt
	response.write "삭제건수=" + CStr(resultCnt) + "<br>"
	Call debugRwite("stepA-5")
    
	''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
	Call debugRwite("stepA-6") ''여기가 느림.

	''itemcouponidx에 등록된 상품의 모든 쿠폰상태 점검후 Update
	sqlStr = "update [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " set itemcouponyn='Y'"
	sqlStr = sqlStr + " ,itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " ,itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " ,curritemcouponidx=T.itemcouponidx"
	IF (Not notUpdate) then
	    sqlStr = sqlStr + " ,lastupdate=getdate()"
    end if
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue, d.itemid "
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_coupon_master m,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item_coupon_detail d"
	sqlStr = sqlStr + " 	where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " 	and m.openstate='7'"
	sqlStr = sqlStr + " 	and d.itemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and m.itemcouponstartdate<=getdate()"
	sqlStr = sqlStr + " 	and m.itemcouponexpiredate>=getdate()"
	sqlStr = sqlStr + " ) T "
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"
	sqlStr = sqlStr + " and Not ("
	sqlStr = sqlStr + " 		 	[db_item].[dbo].tbl_item.itemcouponyn='Y'"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.curritemcouponidx=T.itemcouponidx"
	sqlStr = sqlStr + "			)"

	'response.write sqlStr + "<br>"
	dbget.Execute sqlStr, resultCnt
	Call debugRwite("stepA-7")
    response.write "수정건수=" + CStr(resultCnt)
    
    ''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
	Call debugRwite("stepA-8")
end function

'### 수정 로그 저장
Sub AddSCMChangeLog(couponIdx,logMessage)
	Dim strSql
	if logMessage<>"" then
		strSql = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		strSql = strSql & "VALUES('" & session("ssBctId") & "', 'itemCoupon', '" & couponIdx & "', '" & requestCheckVar(Request("menupos"),9) & "', "
		strSql = strSql & "'" & logMessage & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(strSql)
	end if
End Sub

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim itemcouponidx
dim couponGubun
dim itemcoupontype
dim itemcouponvalue
dim itemcouponstartdate
dim itemcouponexpiredate
dim itemcouponname
dim itemcouponimage
dim itemcouponexplain
dim applyitemcount
dim openstate
dim margintype
dim defaultmargin
dim mode
dim IsEditMode
dim sqlstr,i
dim buf
dim itemidarr, couponbuypricearr, couponsellcasharr, makerid, sailyn
dim ErrStr

dim sType, addSql, itemid, itemname, sellyn, usingyn, danjongyn, disp, couponyn, minmargin, itemcostup, itemcostdown
dim limityn, mwdiv, cdl, cdm, cds, deliverytype, coupontype, groupId
dim itemcouponidxarr
dim exceptnotepmapitem

itemcouponidx      	= requestCheckVar(request("itemcouponidx"),9)
couponGubun         = requestCheckVar(request("couponGubun"),9)
itemcoupontype      = requestCheckVar(request("itemcoupontype"),9)
itemcouponvalue     = requestCheckVar(request("itemcouponvalue"),9)
itemcouponstartdate = request("itemcouponstartdate") + " " + request("itemcouponstartdate2")
itemcouponexpiredate= request("itemcouponexpiredate") + " " + request("itemcouponexpiredate2")
itemcouponname      = html2Db(request("itemcouponname"))
itemcouponimage     = request("itemcouponimage")
applyitemcount      = request("applyitemcount")
openstate         	= request("openstate")
margintype          = request("margintype")
defaultmargin		= request("defaultmargin")
mode 				= request("mode")
itemidarr			= request("itemidarr")
couponbuypricearr	= request("couponbuypricearr")
couponsellcasharr   = request("couponsellcasharr")
itemcouponexplain	= html2Db(request("itemcouponexplain"))
makerid				= request("makerid")
sailyn				= request("sailyn")
sType               = request("sType")

addSql              = request("addSql")
itemid              = request("itemid")
itemname            = request("itemname")
sellyn              = request("sellyn")
usingyn             = request("usingyn")
danjongyn           = request("danjongyn")
limityn             = request("limityn")
mwdiv               = request("mwdiv")
cdl                 = request("cdl")
cdm                 = request("cdm")
cds                 = request("cds")
deliverytype        = request("deliverytype")
coupontype			= requestCheckVar(request("coupontype"),1)
itemcouponidxarr    = request("itemcouponidxarr")
disp                = requestCheckVar(request("disp"),30)
couponyn            = requestCheckVar(request("couponyn"),10)
minmargin           = requestCheckVar(request("minmargin"),10)
itemcostup          = requestCheckVar(request("itemcostup"),10)
itemcostdown        = requestCheckVar(request("itemcostdown"),10)
exceptnotepmapitem	= requestCheckVar(request("exceptnotepmapitem"),10)
groupId				= requestCheckVar(request("groupId"),8)

if itemcouponidx="" then itemcouponidx="0"
if defaultmargin="" then defaultmargin=0
if coupontype="" then coupontype="N"
if (itemcouponidx<>"0") then
	IsEditMode = true
else
	IsEditMode = false
end if

if mode="couponmaster" then
	on Error Resume Next
		buf = CDate(itemcouponstartdate)
		if Err then
			response.Write "<script>alert('발급시작일 오류-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	on Error Goto 0

	on Error Resume Next
		buf = CDate(itemcouponexpiredate)
		if Err then
			response.Write "<script>alert('발급종료일 오류-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	on Error Goto 0

	if (itemcoupontype="1") then
		if (itemcouponvalue>=100) or (itemcouponvalue<1) then
			response.Write "<script>alert('할인쿠폰은 1~99% 사이 값만 가능합니다.')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	elseif (itemcoupontype="2") Then
		If session("ssBctId")="bborami" Then '// 디스전 이벤트 쿠폰 입력때문에 임보람 대리만 제한 풀어둠 단, 날짜제한 걸어둠..
			If Left(Now(), 10) >= "2015-05-12" And Left(Now(), 10) < "2015-05-27" Then
			Else
				if (itemcouponvalue<100) or (itemcouponvalue>=150001) then  ''100000 => 150000 정승훈 요청
					response.Write "<script>alert('할인쿠폰은 1~150000 사이 값만 가능합니다.')</script>"
					response.Write "<script>history.back()</script>"
					dbget.close()	:	response.End
				end If
			End If
		Else
			if (itemcouponvalue<100) or (itemcouponvalue>=300001) then  ''150000 => 300000 김광민 요청 150708
				response.Write "<script>alert('할인쿠폰은 1~300000 사이 값만 가능합니다.')</script>"
				response.Write "<script>history.back()</script>"
				dbget.close()	:	response.End
			end If
		End If
	elseif (itemcoupontype="3") then
		if (Cint(itemcouponvalue)<>Cint(getDefaultBeasongPayByDate(now()))) then
			response.Write "<script>alert('무료배송 할인쿠폰은 " + Cstr(getDefaultBeasongPayByDate(now())) + " 값만 가능합니다.')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	else
		response.Write "<script>alert('쿠폰타입이 지정되지 않았습니다.')</script>"
		response.Write "<script>history.back()</script>"
		dbget.close()	:	response.End
	end if


	if (IsEditMode) then
		''수정
		dim orgDefaultMargin ,orgDefaultMargintype
		sqlstr = "SELECT defaultmargin,margintype FROM db_item.dbo.tbl_item_coupon_master "
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		IF not rsget.eof Then
			orgDefaultMargin = rsget("defaultmargin")
			orgDefaultMargintype = rsget("margintype")
		End IF

		rsget.close

		sqlstr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set itemcoupontype='" + itemcoupontype + "'" + VbCrlf
		sqlstr = sqlstr + " ,couponGubun='" + couponGubun + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponvalue=" + CStr(itemcouponvalue) + VbCrlf
		sqlstr = sqlstr + " ,itemcouponstartdate='" + itemcouponstartdate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexpiredate='" + itemcouponexpiredate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponname='" + itemcouponname + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexplain='" + itemcouponexplain + "'" + VbCrlf
		sqlstr = sqlstr + " ,margintype='" + margintype + "'" + VbCrlf
		sqlstr = sqlstr + " ,defaultmargin='" + defaultmargin + "'" + VbCrlf
		sqlstr = sqlstr + " ,coupontype='" + coupontype + "'" + VbCrlf
		sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
		sqlStr = sqlStr + ", itemcouponimage='" + itemcouponimage + "'" + VbCrlf
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		dbget.Execute sqlStr

        ''2차서버 2018/06/19 
	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	    dblogicsget.Execute sqlStr
	    
		'마진 설정 변경시 대상 상품 전체 변경
		IF (Cint(orgDefaultMargin) <> Cint(defaultmargin)) or (CStr(orgDefaultMargintype)<>CStr(margintype)) Then
				sqlStr =" UPDATE [db_item].[dbo].tbl_item_coupon_detail  "& VbCRLF
				sqlStr = sqlStr& " SET couponbuyprice="& VbCRLF
				SELECT Case margintype
					Case "00"  	''상품개별설정 - 매입가 0 인경우 원매입가
						sqlStr = sqlStr & " 0 " & VbCrlf
					Case "10"	''텐바이텐부담 - 원매입가
						sqlStr = sqlStr & " 0 " & VbCrlf
					Case "20"	''직접설정 : 추가 [2008-09-23]
						if itemcoupontype="1" then			''할인율
							sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
						elseif itemcoupontype="2" then   	''금액
							sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "30"	''동일마진 - 현재마진 : 추가 [2008-09-23]
						if itemcoupontype="1" then			''할인율
							sqlStr = sqlStr & " convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)"
						elseif itemcoupontype="2" then   	''금액
							sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "50"	''반반부담
						if itemcoupontype="1" then			''할인율
							sqlStr = sqlStr & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)"
						elseif itemcoupontype="2" then   	''금액
							sqlStr = sqlStr & " i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "60"	''업체부담 - 매입가 조정
						if itemcoupontype="1" then			''할인율
							sqlStr = sqlStr & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) + ")"
						elseif itemcoupontype="2" then   	''금액
							sqlStr = sqlStr & " i.buycash - " & CStr(itemcouponvalue)
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
			        Case "80"   ''무료배송쿠폰 -500
			            sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
					Case "90"	''20%전체행사 - 매입인경우 원매입가.
						if itemcoupontype="1" then			''할인율
							sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5) end "
						elseif itemcoupontype="2" then   	''금액
							sqlStr = sqlStr & " case when i.mwdiv='M' 0 else i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)  end "
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case else
						sqlStr = sqlStr & " 0 " & VbCrlf
				End SELECT
						sqlStr = sqlStr & " , couponmargin=" 
				SELECT Case margintype
					Case "00"  	''상품개별설정 - 매입가 0 인경우 원매입가 
            		    sqlStr = sqlStr & " 0" & VbCrlf
            		Case "10"	''텐바이텐부담 - 원매입가 
                        sqlStr = sqlStr & " 0" & VbCrlf
            		Case "20"	''직접설정 : 추가 [2008-09-23] 
            			if itemcoupontype="1" then			''할인율 
                            sqlStr = sqlStr & " (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
            			elseif itemcoupontype="2" then   	''금액 
            				sqlStr = sqlStr & " (((i.sellcash-"&Cstr(itemcouponvalue)&") -(convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "30"	''동일마진 - 현재마진 : 추가 [2008-09-23]
            			if itemcoupontype="1" then			''할인율 
            				sqlStr = sqlStr & "  (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''금액 
            				sqlStr = sqlStr & "  (((i.sellcash-"&Cstr(itemcouponvalue)&") - (convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "50"	''반반부담
            			if itemcoupontype="1" then			''할인율 
            				sqlStr = sqlStr  & "  (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - ( i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''금액 
            				sqlStr = sqlStr & "  (((i.sellcash-"&Cstr(itemcouponvalue)&")- ( i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "60"	''업체부담 - 매입가 조정
            			if itemcoupontype="1" then			''할인율 
            				sqlStr = sqlStr  & " (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''금액 
            				sqlStr = sqlStr  & " (((i.sellcash-"&Cstr(itemcouponvalue)&") - (i.buycash - " & CStr(itemcouponvalue) &"))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
                    Case "80"   ''무료배송쿠폰 -500 
                            sqlStr = sqlStr  & " case when i.mwdiv='M' then 0 else ((i.sellcash- (i.buycash - 500))/i.sellcash)*100 end "
            		Case "90"	''20%전체행사 - 매입인경우 원매입가.
            			if itemcoupontype="1" then			''할인율 
            				sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash))-(i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 end "
            			elseif itemcoupontype="2" then   	''금액 
            				sqlStr = sqlStr & " case when i.mwdiv='M' 0 else (((i.sellcash-"&Cstr(itemcouponvalue)&")-(i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100  end "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            
            		Case else 
            			sqlStr = sqlStr & " 0 " & VbCrlf
				End SELECT
				sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_coupon_detail d " & VbCrlf
				sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_item i "
				sqlStr = sqlStr & " 	on d.itemid = i.itemid "
				sqlStr = sqlStr & " WHERE d.itemcouponidx=" & CStr(itemcouponidx)
          
        '  response.write sqlStr
       ' response.end
				dbget.Execute sqlStr
				
				''2차서버 2018/06/19 
        	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
        	    dblogicsget.Execute sqlStr
		End IF

	else
		''신규 등록
		sqlStr = "select * from [db_item].[dbo].tbl_item_coupon_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew

		rsget("itemcoupontype") = itemcoupontype
		rsget("couponGubun") = couponGubun
		rsget("itemcouponvalue") = itemcouponvalue
		rsget("itemcouponstartdate") = itemcouponstartdate
		rsget("itemcouponexpiredate") = itemcouponexpiredate
		rsget("itemcouponname") = itemcouponname
		rsget("itemcouponexplain") = itemcouponexplain

		rsget("openstate") = "0"
		rsget("margintype") = margintype
		rsget("defaultmargin")	= defaultmargin
		rsget("reguserid") = session("ssBctId")
		rsget("coupontype") = coupontype

		rsget.update
			itemcouponidx = rsget("itemcouponidx")
		rsget.close

	end if
elseif mode="I" then
    '' 추가 팝업창에서 넘어 올 경우.
	ErrStr = ""

	''마진타입 가져오기
	margintype = "00"

	sqlStr = "select top 1 margintype, itemcoupontype, itemcouponvalue,couponGubun,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		margintype = rsget("margintype")
		itemcoupontype = rsget("itemcoupontype")
		itemcouponvalue = rsget("itemcouponvalue")
		itemcouponstartdate = rsget("itemcouponstartdate")
		itemcouponexpiredate = rsget("itemcouponexpiredate")
		couponGubun = rsget("couponGubun")  ''일반/네이버/지정인등.
	end if
	rsget.close

	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	'' 무료배송 쿠폰일경우, 업체상품 및 텐배무료배송 기준금액 초과 상품 안내
	if itemcoupontype=3 then
		sqlStr = "Select top 100 itemid, mwdiv, sellcash " & vbCRLF
		sqlStr = sqlStr & " from db_item.dbo.tbl_item " & vbCRLF
		sqlStr = sqlStr & " Where itemid in (" & itemidarr & ")"
		
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.Eof then
			do until rsget.Eof
				if rsget("mwdiv")="U" then ErrStr = ErrStr + "-업체배송 상품 (상품번호 : " + CStr(rsget("itemid")) + ") 등록불가 \n"
				if rsget("mwdiv")<>"U" and rsget("sellcash")>=30000 then ErrStr = ErrStr + "- 무료배송 상품 (상품번호 : " + CStr(rsget("itemid")) + ") 등록불가 \n"
				rsget.moveNext
			loop

			if ErrStr<>"" then
				response.write "<script language=javascript>alert('배송료할인 쿠폰에는\n\n" + ErrStr + "');</script>"
				response.End
			end if
		end if
		rsget.close
	end if
	
	Call debugRwite("step1")
    ''검색한 전체 상품인 경우.. 검색된 모든 내용 insert  처리
    addSql = ""
    IF (sType="all") THEN

         '// 추가 쿼리
		
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemid <> "") then
			itemid = trim(itemid)
			itemid = replace(itemid,chr(13),"")
			itemid = replace(itemid,chr(10),",")
			if Right(itemid,1)="," then itemid=Left(itemid,Len(itemid)-1)

            addSql = addSql & " and i.itemid in (" + itemid + ")"
        end if

        ''if (itemname <> "") then
        ''    addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
        ''end if

        if (sellyn="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( sellyn="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (sellyn <> "") then
            addSql = addSql & " and i.sellyn='" + sellyn + "'"
        end if
        
        if (usingyn <> "") then
            addSql = addSql & " and i.isusing='" + usingyn + "'"
        end if

        if danjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif danjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif danjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
        end if
       
        if mwdiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif mwdiv<>"" then
            addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
        end if
        
		if limityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif limityn<>"" then
            addSql = addSql + " and i.limityn='" + limityn + "'"
        end if

        if cdl<>"" then
            addSql = addSql + " and i.cate_large='" + cdl + "'"
        end if

        if cdm<>"" then
            addSql = addSql + " and i.cate_mid='" + cdm + "'"
        end if

        if cds<>"" then
            addSql = addSql + " and i.cate_small='" + cds + "'"
        end if

        if disp<>"" then
		    if LEN(disp)>3 then
		         addSql = addSql + " and i.dispcate1='"&LEFT(disp,3)&"'" ''2015/03/27추가
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item WITH(NOLOCK) where catecode like '" + disp + "%' and isDefault='y') "
		end if
		
		if couponyn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + couponyn + "'"
        end if

        if sailyn<>"" then
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if

        if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if

        If minmargin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.sellcash <> 0 and ((1-(i.buycash/i.sellcash))*100) >= " & minmargin & " "
        End If
        
        if (itemcostup<>"") then
            addSql = addSql & " and i.sellcash>="&itemcostup&""&vbCRLF
        end if
        
        if (itemcostdown<>"") then
            addSql = addSql & " and i.sellcash<="&itemcostdown&""&vbCRLF
        end if

		if (groupId<>"") then
            addSql = addSql & " and Exists ( "
            addSql = addSql & " 	select 1 "
            addSql = addSql & " 	from db_partner.dbo.tbl_partner as sp "
            addSql = addSql & " 		join db_user.dbo.tbl_user_c as sc "
            addSql = addSql & " 			on sp.id=sc.userid "
            addSql = addSql & " 	where sp.id=i.makerid "
            addSql = addSql & " 		and sp.isusing='Y' and sc.isusing='Y' "
            addSql = addSql & " 		and sc.userdiv='02' "
            addSql = addSql & " 		and sp.groupid='" & groupId & "' "
            addSql = addSql & " ) "
		end if

        ''EP 제외조건
        addSql = addSql & " and i.itemdiv<>'21'"
        addSql = addSql & " and i.makerid not in (select makerid from db_temp.dbo.tbl_EpShop_not_in_makerid WITH(NOLOCK) where mallgubun='naverep' and isusing='N')"
        addSql = addSql & " and i.itemid not in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid WITH(NOLOCK) Where mallgubun='naverep' AND isusing = 'Y')"
        ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_Mapping_item)"
        ''addSql = addSql & " and Not Exists(select 1 from db_temp.dbo.tbl_naver_item_map nn where nn.serviceyn='y' and nn.tenitemid=i.itemid)"   ''tbl_nvshop_mapItem 으로변경 
        ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_RecentSell_item where (sellNDays>=6 or sell1Days>=2))"  ''최근 판매내역 N개이상 제외 (주석처리 2018/07/19)
        
        ''2018/07/18
        addSql = addSql & " and i.makerid not in ( select makerid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_Brand WITH(NOLOCK) where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"
        addSql = addSql & " and i.itemid not in ( select itemid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_item WITH(NOLOCK) where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"
        
		''2019/11/04 조건추가.
        addSql = addSql & " and Not Exists(select 1 from [db_temp].dbo.[tbl_Epshop_fixedPrice] fx WITH(NOLOCK) where fx.itemid=i.itemid)"
		''조건 분기처리. 희란님 요청 2019/11/04
        if (exceptnotepmapitem="") then
			addSql = addSql & " and Not Exists(select 1 from [db_etcmall].dbo.[tbl_nvshop_mapItem] nn WITH(NOLOCK) where nn.itemid=i.itemid)" 
		end if

        ''등록예정쿠폰제외
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid"
        addSql = addSql & "     from [db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) "
        addSql = addSql & "         Join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
        addSql = addSql & "         on m.itemcouponidx=d.itemcouponidx"
        addSql = addSql & "         and m.openstate<9"
        addSql = addSql & "     where m.itemcouponexpiredate>getdate()"
		addSql = addSql + " 	and NOT ("
		addSql = addSql + " 		(m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
		addSql = addSql + " 		or "
		addSql = addSql + " 		(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "')"
        ' addSql = addSql & "     (m.itemcouponstartdate<='"&itemcouponstartdate&"' and m.itemcouponexpiredate>'"&itemcouponstartdate&"')"
        ' addSql = addSql & "     or"
        ' addSql = addSql & "     (m.itemcouponstartdate<='"&itemcouponexpiredate&"' and m.itemcouponexpiredate>'"&itemcouponexpiredate&"')"
        addSql = addSql & "     )"
		'if (couponGubun="V") then  ''중복쿠폰 체크방식 변경 //2019/03/25 => 검색후 대상 입력하는 케이스는 이방식을 타지 않는다.
		'	addSql = addSql + " and m.couponGubun='V'"
		'else
		'	addSql = addSql + " and m.couponGubun<>'V'"
		'end if
        addSql = addSql & " )"

        if (addSql="") then
            addSql = "select i.itemid from [db_item].[dbo].tbl_item i WITH(NOLOCK) where 1=0 "
        else
            addSql = "select i.itemid from [db_item].[dbo].tbl_item i WITH(NOLOCK) where i.itemid<>0 " & addSql
        end if
         
        '' counting 검증
        dim iCountQuery, paraitemcount, isubcnt : isubcnt=0
        iCountQuery = replace(addSql,"select i.itemid","select count(*) cnt")

        rsget.CursorLocation = adUseClient
        rsget.Open iCountQuery, dbget, adOpenForwardOnly, adLockReadOnly
	    if not rsget.Eof then
	        isubcnt = rsget("cnt")
	    end if
	    rsget.close
	    
	    paraitemcount = request("itemcount")
	    
	    if (CStr(paraitemcount)<>CStr(isubcnt)) then
	        
	        response.write "<script>alert('수량 오류 :"&paraitemcount&":"&isubcnt&"');</script>"
	        rw addSql
	        dbget.Close() : response.end
	    end if
    ELSE
    	addSql = trim(itemidarr)
	END IF
	Call debugRwite("step2")

	'' 다른 쿠폰에 상품이 등록되어 있을경우 체크
	sqlStr = " select top 100 m.itemcouponidx, d.itemid from"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) "
	sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
	sqlStr = sqlStr + " on m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " where m.itemcouponidx<>" + CStr(itemcouponidx)
	sqlStr = sqlStr + " and m.openstate<9"			''발급종료인것 제외
	'sqlStr = sqlStr + " and m.couponGubun<>'P'"		''지정인발급쿠폰은 제외 (20140617; 허진원) , 중복쿠폰 불가능(2018.01.22)
	' sqlStr = sqlStr + " and ( "
	' sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponstartdate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponstartdate) + "')"
	' sqlStr = sqlStr + " 	or "
	' sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponexpiredate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponexpiredate) + "')"
	' sqlStr = sqlStr + " 	)"
	sqlStr = sqlStr + " and m.itemcouponexpiredate>getdate()"
	sqlStr = sqlStr + " and NOT ( "
	sqlStr = sqlStr + " 	(m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 	or "
	sqlStr = sqlStr + " 	(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "')"
	sqlStr = sqlStr + " 	)"

	if (sType<>"all") then  ''검색후 대상 입력하는 케이스는 이방식을 타지 않는다. //2019/03/29
		if (couponGubun="V") then  ''중복쿠폰 체크방식 변경 //2019/03/25
			sqlStr = sqlStr + " and m.couponGubun='V'"
		else
			sqlStr = sqlStr + " and m.couponGubun not in ('V','P','T')"  ''P,T추가(secret==P 상관없음.) 2019/06/11
		end if
	end if
	sqlStr = sqlStr + " and d.itemid in (" + addSql + ")"  + VbCrlf

	

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		do until rsget.Eof
			ErrStr = ErrStr + "쿠폰번호 : " + CStr(rsget("itemcouponidx")) + " - 상품번호 : " + CStr(rsget("itemid")) + " 사용중 \n"
			rsget.moveNext
		loop
	end if
	rsget.close

	Call debugRwite("step3")
	'' 기존 목록에 없는 상품만 추가.
	sqlStr = "insert into [db_item].[dbo].tbl_item_coupon_detail " & VbCrlf
	sqlStr = sqlStr & " (itemcouponidx, itemid, couponbuyprice, couponmargin)" & VbCrlf
	sqlStr = sqlStr & " select "& CStr(itemcouponidx) & "," & VbCrlf
	sqlStr = sqlStr & " i.itemid, " & VbCrlf
	Select Case margintype
		Case "00"  	''상품개별설정 - 매입가 0 인경우 원매입가
			sqlStr = sqlStr & " 0 " & VbCrlf
		'Case "10"	''텐바이텐부담 - 매입가 조정x
		'	if itemcoupontype="1" then			''할인율
		'		sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + ")"
		'	elseif itemcoupontype="2" then   	''금액
		'		sqlStr = sqlStr + " i.buycash - " + CStr(itemcouponvalue)
		'	else
		'		sqlStr = sqlStr + " 0 " + VbCrlf
		'	end if
		    sqlStr = sqlStr & ", 0" & VbCrlf
		Case "10"	''텐바이텐부담 - 원매입가
			sqlStr = sqlStr & " 0 " & VbCrlf
            sqlStr = sqlStr & ", 0" & VbCrlf
		Case "20"	''직접설정 : 추가 [2008-09-23]
		 
			if itemcoupontype="1" then			''할인율
				sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
                sqlStr = sqlStr & ", ( ( (i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")) )/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''금액
				sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
				sqlStr = sqlStr & " ,(((i.sellcash-"&Cstr(itemcouponvalue)&") -(convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
			else
				sqlStr = sqlStr & " 0 " & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "30"	''동일마진 - 현재마진 : 추가 [2008-09-23]
			if itemcoupontype="1" then			''할인율
				sqlStr = sqlStr & " convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)"
				sqlStr = sqlStr & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
			elseif itemcoupontype="2" then   	''금액
				sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)"
				sqlStr = sqlStr & " , (((i.sellcash-"&Cstr(itemcouponvalue)&") - (convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
			else
				sqlStr = sqlStr & " 0 " & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "50"	''반반부담
			if itemcoupontype="1" then			''할인율
				sqlStr = sqlStr  & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)"
				sqlStr = sqlStr  & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - ( i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''금액
				sqlStr = sqlStr & " i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)"
				sqlStr = sqlStr & " , (((i.sellcash-"&Cstr(itemcouponvalue)&")- ( i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "60"	''업체부담 - 매입가 조정
			if itemcoupontype="1" then			''할인율
				sqlStr = sqlStr  & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")"
				sqlStr = sqlStr  & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''금액
				sqlStr = sqlStr  & " i.buycash - " & CStr(itemcouponvalue)
				sqlStr = sqlStr  & "  , (((i.sellcash-"&Cstr(itemcouponvalue)&") - (i.buycash - " & CStr(itemcouponvalue) &"))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
        Case "80"   ''무료배송쿠폰 -500
                sqlStr = sqlStr  & " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
                sqlStr = sqlStr  & ", case when i.mwdiv='M' then 0 else ((i.sellcash- (i.buycash - 500))/i.sellcash)*100 end "
		Case "90"	''20%전체행사 - 매입인경우 원매입가.
			if itemcoupontype="1" then			''할인율
				sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5) end "
				sqlStr = sqlStr & ", case when i.mwdiv='M' then 0 else (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash))-(i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 end "
			elseif itemcoupontype="2" then   	''금액
				sqlStr = sqlStr & " case when i.mwdiv='M' 0 else i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)  end "
				sqlStr = sqlStr & ", case when i.mwdiv='M' 0 else (((i.sellcash-"&Cstr(itemcouponvalue)&")-(i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100  end "
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if

		Case else
			sqlStr = sqlStr  & " 0 "  & VbCrlf
			sqlStr = sqlStr & " ,0 " & VbCrlf
	end Select

	sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i WITH(NOLOCK) " &VbCrlf
	sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) " & VbCrlf
	sqlStr = sqlStr & " 	on d.itemcouponidx=" & CStr(itemcouponidx) & VbCrlf
	sqlStr = sqlStr & " 	and d.itemid=i.itemid" & VbCrlf
	sqlStr = sqlStr & " where i.itemid in (" & addSql & ")"  & VbCrlf
	sqlStr = sqlStr & " 	and d.itemid is null"
	sqlStr = sqlStr + "		and i.itemdiv<>'21' "  ''딜상품 제외
	sqlStr = sqlStr & " 	and i.itemid not in ("
	sqlStr = sqlStr & " 		select distinct d.itemid from"
	sqlStr = sqlStr & " 			[db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) ,"
	sqlStr = sqlStr & " 			[db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
	sqlStr = sqlStr & " 		where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr & " 			and m.itemcouponidx<>" & CStr(itemcouponidx)
	sqlStr = sqlStr & " 			and m.openstate<9"  ''발급종료인것 제외
	sqlStr = sqlStr + " 			and m.itemcouponexpiredate>getdate()"
	sqlStr = sqlStr + " 			and NOT ( "
	sqlStr = sqlStr + " 				((m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 				or "
	sqlStr = sqlStr + " 				(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "'))"
	sqlStr = sqlStr + " 			)"
	if (sType<>"all") then  ''검색후 대상 입력하는 케이스는 이방식을 타지 않는다. //2019/03/29
		if (couponGubun="V") then  ''중복쿠폰 체크방식 변경 //2019/03/25
			sqlStr = sqlStr + " and m.couponGubun='V'"
		else
			sqlStr = sqlStr + " and m.couponGubun not in ('V','P','T') "
		end if
	end if
	sqlStr = sqlStr & " 	and d.itemid in (" & addSql & ")"  & VbCrlf
	sqlStr = sqlStr & " ) "
 
	dbget.CommandTimeout = 150	'5분
	dbget.Execute sqlStr
	Call debugRwite("step4")
	
	Call AplyToItem(itemcouponidx,false)
	Call debugRwite("step5")
	''적용상품수.
	AplyItemCountUpdate itemcouponidx
	
	Call debugRwite("step6")
	if Not(itemid="" and itemidarr="") then
		Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>상품추가 : " & itemid & itemidarr)
	end if
elseif mode="delcouponitemarr" then
	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	sqlStr = "delete from [db_item].[dbo].tbl_item_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " and itemid in (" + itemidarr + ")"  + VbCrlf

	dbget.Execute sqlStr

    ''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
    ''삭제인경우 2차서버에 lastupdate를 쳐주자.. Naver EP관련.
    sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
    sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
    sqlStr = sqlStr + " where itemid in (" + itemidarr + ")"  + VbCrlf
    dblogicsget.Execute sqlStr
    
	''삭제된 쿠폰 상품테이블에서 쿠폰 여부 N 로 변경
	Call AplyToItem(itemcouponidx,false)

	''적용상품수.
	AplyItemCountUpdate itemcouponidx

	if itemidarr<>"" then
		Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>상품삭제 : " & itemidarr)
	end if
elseif mode="delBrandAll" then
	'// 브랜드 상품 일괄 삭제
	if makerid<>"" then
		sqlStr = "delete from cd " + VbCrlf
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_coupon_detail as cd with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 	join [db_item].[dbo].tbl_item as i with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 		on cd.itemid=i.itemid " + VbCrlf
		sqlStr = sqlStr & " where cd.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
		sqlStr = sqlStr & " 	and i.makerid='" + makerid + "' " + VbCrlf
		dbget.Execute sqlStr

		''삭제인경우 2차서버에에서 먼저 lastupdate를 쳐주자.. Naver EP관련.
		sqlStr = "update i "+ VbCrlf
		sqlStr = sqlStr + " set i.lastupdate =getdate()"+ VbCrlf
		sqlStr = sqlStr & " from [db_AppWish].[dbo].tbl_item as i with(noLock) " + VbCrlf
		sqlStr = sqlStr & "		join [db_AppWish].[dbo].tbl_item_coupon_detail as cd with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 		on i.itemid=cd.itemid " + VbCrlf
		sqlStr = sqlStr & " where cd.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
		sqlStr = sqlStr & " 	and i.makerid='" + makerid + "' " + VbCrlf
		dblogicsget.Execute sqlStr

		''2차서버 2018/06/19 
		sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
		dblogicsget.Execute sqlStr


		''삭제된 쿠폰 상품테이블에서 쿠폰 여부 N 로 변경
		Call AplyToItem(itemcouponidx,false)

		''적용상품수.
		AplyItemCountUpdate itemcouponidx

		Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>브랜드상품삭제 : " & makerid)
	end if

elseif mode="delcouponitemmulti" then  ''2018/05/17
    dim midxArrQue
    itemcouponidxarr = split(itemcouponidxarr,",")
    itemidarr        = split(itemidarr,",")
    
    if (Lbound(itemcouponidxarr)<>Lbound(itemidarr)) or (Ubound(itemcouponidxarr)<>Ubound(itemidarr)) then
        response.Write "<script>alert('param 오류')</script>"
		response.Write "<script>history.back()</script>"
		dbget.close()	:	response.End
    end if
    
    for i=Lbound(itemcouponidxarr) to Ubound(itemcouponidxarr)
        if (itemcouponidxarr(i)<>"") and (itemidarr(i)<>"") then
            ''rw itemcouponidxarr(i)&","&itemidarr(i)
            
            sqlStr = "delete from [db_item].[dbo].tbl_item_coupon_detail" + VbCrlf
        	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidxarr(i)) + VbCrlf
        	sqlStr = sqlStr + " and itemid in (" + itemidarr(i) + ")"  + VbCrlf
        
        	dbget.Execute sqlStr
        	
        	''2차서버 2018/06/19 
        	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
            dblogicsget.Execute sqlStr
            ''삭제인경우 2차서버에 lastupdate를 쳐주자.. Naver EP관련.
            sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
            sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
            sqlStr = sqlStr + " where itemid in (" + itemidarr(i) + ")"  + VbCrlf
            dblogicsget.Execute sqlStr
    
        	if Not(InStr(midxArrQue,itemcouponidxarr(i)&",")>0) then
        	    midxArrQue = midxArrQue&itemcouponidxarr(i)&","
        	end if

			Call AddSCMChangeLog(itemcouponidxarr(i), "- 상품쿠폰>상품삭제 : " & itemidarr(i))
        end if
    next
    
    midxArrQue = split(midxArrQue,",")
    for i=Lbound(midxArrQue) to Ubound(midxArrQue)
        if (midxArrQue(i)<>"") then
            ''삭제된 쿠폰 상품테이블에서 쿠폰 여부 N 로 변경
        	Call AplyToItem(midxArrQue(i),false)
        
        	''적용상품수.
        	AplyItemCountUpdate midxArrQue(i)


        end if
    next
    
elseif mode="modicouponitemarr" then
	itemidarr = trim(itemidarr)
	couponbuypricearr  = trim(couponbuypricearr)
    couponsellcasharr = trim(couponsellcasharr)
    
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)
	if Right(couponbuypricearr,1)="," then couponbuypricearr=Left(couponbuypricearr,Len(couponbuypricearr)-1)
	if Right(couponsellcasharr,1)="," then couponsellcasharr=Left(couponsellcasharr,Len(couponsellcasharr)-1)

	itemidarr = split(itemidarr,",")
	couponbuypricearr = split(couponbuypricearr,",")
    couponsellcasharr = split(couponsellcasharr,",")
    
	for i=LBound(itemidarr) to UBound(itemidarr)
		if trim(itemidarr(i))<>"" then
			sqlStr = "update D" + VbCrlf
			sqlStr = sqlStr + " set couponbuyprice=" + CStr(couponbuypricearr(i)) + VbCrlf
			if (TRIM(couponbuypricearr(i))="0") or (TRIM(couponsellcasharr(i))="0") or (TRIM(couponsellcasharr(i))="") then
			    sqlStr = sqlStr + " ,couponmargin=0" + VbCrlf
			else
			    sqlStr = sqlStr + " ,couponmargin=(1-" +CStr(couponbuypricearr(i))+"*1.0/"+CStr(couponsellcasharr(i))+")*100"
			end if
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_detail D" + VbCrlf
			sqlStr = sqlStr + " where D.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
			sqlStr = sqlStr + " and D.itemid=" + CStr(itemidarr(i)) + VbCrlf

			dbget.Execute sqlStr
			
			''2차서버 2018/06/19 
        	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
            dblogicsget.Execute sqlStr
            ''2차서버 lastupdate for Naver coupon
            sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
            sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
            sqlStr = sqlStr + " where itemid in (" + itemidarr(i) + ")"  + VbCrlf
            dblogicsget.Execute sqlStr
		end if
	next

	Call AplyToItem(itemcouponidx,false)

	''적용상품수.
	AplyItemCountUpdate itemcouponidx
elseif mode="opencoupon" Then

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='7'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
'response.write sqlStr
	Call AplyToItem(itemcouponidx,true)
	
	''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>쿠폰오픈")

elseif mode="reservecoupon" Then

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='6'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr

    ''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>쿠폰예약")

elseif mode="closecoupon" Then

    dim MayExpireDt
    MayExpireDt = Left(CStr(DateAdd("d",-1,Now())),10) & " 23:59:59"

    ''response.write MayExpireDt

    ''기 발급 된 쿠폰 Expire
    sqlStr = "update [db_item].[dbo].tbl_user_item_coupon" + VbCrlf
    sqlStr = sqlStr + " set itemcouponexpiredate='" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
    sqlStr = sqlStr + " and itemcouponexpiredate>'" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " and usedyn='N'" + VbCrlf

    dbget.Execute sqlStr
    

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='9'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
	
	''2차서버 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AplyToItem(itemcouponidx,true)

	Call AddSCMChangeLog(itemcouponidx, "- 상품쿠폰>쿠폰종료")

elseif mode="imageupload" Then
		''수정

		sqlstr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set itemcouponimage='" + itemcouponimage + "'" + VbCrlf
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		dbget.Execute sqlStr
end if

%>
<% if (mode="couponmaster") then %>
	<% if (IsEditMode) then %>
	<script language='javascript'>
	alert('수정 되었습니다.');
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	<% else %>
	<script language='javascript'>
	alert('저장 되었습니다. 상품을 등록 해 주세요');
	opener.location.reload();
	window.close();
	//location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	<% end if %>
<% elseif mode="I" then %>
	<script language='javascript'>
	<%
	if ErrStr<>"" then
		ErrStr = ErrStr + "\n\n 쿠폰을 중복으로 발행 할 수 없습니다."
		response.write "alert('" + ErrStr + "')"
	end if
	%>
	alert('상품이 등록 되었습니다.');
	//location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="delcouponitemarr" or mode="delBrandAll" then %>
	<script language='javascript'>
	alert('삭제 되었습니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="delcouponitemmulti" then %>
	<script language='javascript'>
	alert('삭제 되었습니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidtMulti.asp?makerid=<%= makerid %>');
	</script>
<% elseif mode="modicouponitemarr" then %>
	<script language='javascript'>
	alert('수정 되었습니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="opencoupon" then %>
	<script language='javascript'>
	alert('쿠폰이 오픈 되었습니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="reservecoupon" then %>
	<script language='javascript'>
	alert('쿠폰이 오픈이 예약 되었습니다. 매일 0시에 적용됩니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="closecoupon" then %>
	<script language='javascript'>
	alert('쿠폰이 종료 되었습니다.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="imageupload" then %>
	<script language='javascript'>
	alert('수정 되었습니다.');
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% end if %>
<%= "mode=" + mode %>
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->