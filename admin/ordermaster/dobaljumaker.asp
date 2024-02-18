<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 80
%>
<%
'###########################################################
' Description : 출고지시
' History : 이상구 생성
'			2018.03.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim mode, orderserial, sitename, differencekey, workgroup, songjangdiv, baljutype, ems, ups, epostmilitary, extSiteName, isGiftPojang
dim sqlStr,i, iid, dummiserial, obaljudate, errcode, standingorderserial, standingorderitemid, standingorderitemoption, boxType
dim notmatchingstandingorder, startreserveidx, pickingStationCd, pickingStationCdArr, pickingStationCdArrStr
mode        = request("mode")
orderserial = request("orderserial")
sitename    = request("sitename")
workgroup   = request("workgroup")
baljutype   = request("baljutype")
extSiteName   = request("extSiteName")
isGiftPojang   = request("isGiftPojang")
boxType   = request("boxType")

pickingStationCd   = request("pickingStationCd")
pickingStationCdArr 		= pickingStationCd
pickingStationCdArrStr 		= Replace(pickingStationCd, " ", "")
pickingStationCdArr 		= Split(pickingStationCdArr, ",")
if UBound(pickingStationCdArr) > 0 then
    pickingStationCd = "IFC" & (UBound(pickingStationCdArr) + 1)
end if

'// not used
'' ems				= request("ems")
'' epostmilitary	= request("epostmilitary")
'' cn10x10			= request("cn10x10")
'' ecargo			= request("ecargo")
'단품출고설정
if workgroup="N" then baljutype="S"
if workgroup="M" then baljutype="S"
if workgroup="J" then baljutype="S"

''출고 택배사 추가 int
songjangdiv = request("songjangdiv")

dummiserial = orderserial
orderserial = split(orderserial,"|")
sitename    = split(sitename,"|")

if mode="arr" then
	''유효성체크.
	dummiserial = Mid(dummiserial,2,Len(dummiserial))
	dummiserial = replace(dummiserial,"|","','")
	sqlStr = " select top 1 orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " where orderserial in ('" + dummiserial + "')"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		response.write "<script language='javascript'>"
		response.write "alert('" + rsget("orderserial") + " : 이미 출고지시된 주문입니다. \n\n이미 출고지시한 주문건은 중복 출고지시 할 수 없습니다.');"
		response.write "location.replace('" + CStr(refer) + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	'////////////////// 출고이전 정기구독 주문분리 ////////////////		'2019.02.13 한용민
	'/ 주문건 중에 정기구독 상품이 있는지 체크
	sqlStr = "select d.orderserial, so.orgitemid, so.orgitemoption, so.reserveitemid, so.reserveitemoption" & vbcrlf
	sqlStr = sqlStr & " from db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sqlStr = sqlStr & " join db_item.dbo.tbl_item i with (nolock)" & vbcrlf
	sqlStr = sqlStr & " 	on d.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.itemdiv = '75'" & vbcrlf		' 정기구독 상품
	sqlStr = sqlStr & " left join db_item.dbo.tbl_item_standing_item si" & vbcrlf
	sqlStr = sqlStr & " 	on d.itemid=si.orgitemid" & vbcrlf
	sqlStr = sqlStr & " 	and d.itemoption=si.orgitemoption" & vbcrlf
	sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_standing_order] as so" & vbcrlf
	sqlStr = sqlStr & " 	on d.itemid = so.orgitemid and d.itemoption = so.orgitemoption" & vbcrlf
	sqlStr = sqlStr & " 	and si.startreserveidx=so.reserveidx" & vbcrlf
	sqlStr = sqlStr & " where d.itemid not in (0,100)" & vbcrlf
	sqlStr = sqlStr & " and d.cancelyn<>'Y'" & vbcrlf
	sqlStr = sqlStr & " and d.orderserial in ('" & dummiserial & "')" & vbcrlf
	sqlStr = sqlStr & " group by d.orderserial, so.orgitemid, so.orgitemoption, so.reserveitemid, so.reserveitemoption" & vbcrlf

	'response.write sqlStr & "<BR>"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		do until rsget.eof
			' 실제 상품 매칭 안된 주문건
			if isnull(rsget("reserveitemoption")) or rsget("reserveitemoption")="" then
				notmatchingstandingorder = notmatchingstandingorder & rsget("orderserial") & ","
			Else
				' 정기구독주문
				standingorderserial = standingorderserial & rsget("orderserial") & ","
			end if
			rsget.movenext
		loop
	end if
	rsget.Close

	if standingorderserial <> "" then
		standingorderserial = "'" & left(standingorderserial,len(standingorderserial)-1) & "'"
		standingorderserial = replace(standingorderserial,",","','")
	end if
	if notmatchingstandingorder <> "" then
		notmatchingstandingorder = "'" & left(notmatchingstandingorder,len(notmatchingstandingorder)-1) & "'"
		notmatchingstandingorder = replace(notmatchingstandingorder,",","','")

		response.write "<script type='text/javascript'>"
		response.write "	alert('[정기구독] 실제발송 상품 매칭이 지정되어 있지 않습니다.\n해당 주문건은 제외되고 출고지시 됩니다.\n해당 부서에 정기구독 상품매칭 요청하세요.\n\n' + "& notmatchingstandingorder &");"
		response.write "</script>"
	end if

	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			' 정기구독 상품일 경우 에만 쿼리
			if instr(standingorderserial,orderserial(i)) > 0 then
				standingorderitemid = ""
				standingorderitemoption = ""
				startreserveidx = ""

				'/ 정기구독 상품번호와 옵션번호를 받아옴
				sqlStr = "select top 1 d.itemid, d.itemoption, si.startreserveidx" & vbcrlf
				sqlStr = sqlStr & " from db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
				sqlStr = sqlStr & " join db_item.dbo.tbl_item_standing_item si" & vbcrlf
				sqlStr = sqlStr & " 	on d.itemid=si.orgitemid" & vbcrlf
				sqlStr = sqlStr & " 	and d.itemoption=si.orgitemoption" & vbcrlf
				sqlStr = sqlStr & " join db_item.[dbo].[tbl_item_standing_order] as so" & vbcrlf
				sqlStr = sqlStr & " 	on d.itemid = so.orgitemid and d.itemoption = so.orgitemoption" & vbcrlf
				sqlStr = sqlStr & " 	and si.startreserveidx=so.reserveidx" & vbcrlf
				sqlStr = sqlStr & " where d.orderserial in ('" & orderserial(i) & "')" & vbcrlf

				'response.write sqlStr & "<BR>"
				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					standingorderitemid = rsget("itemid")
					standingorderitemoption = rsget("itemoption")
					startreserveidx = rsget("startreserveidx")
				end if
				rsget.Close

				' 히치하이커 정기구독 출고이전 주문분리
				sqlStr = "exec db_order.[dbo].sp_Ten_StandingOrder_make "& standingorderitemid &",'"& standingorderitemoption &"',"& startreserveidx &",'"& orderserial(i) &"','SCM'"

				'response.write sqlStr & "<br>"
				dbget.execute sqlStr
			end if
		end if
	next
	'////////////////// 출고이전 정기구독 주문분리 ////////////////

	sqlStr = " select (count(id) + 1) as differencekey"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljumaster"
	sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21)"

	rsget.Open sqlStr,dbget,1
		differencekey = rsget("differencekey")
	rsget.close

On Error Resume Next
dbget.beginTrans

If Err.Number = 0 Then
        errcode = "001"

    '' 택배사별 출고지시 system으로 수정

	'#######출고지시 마스터###############
    if (Left(pickingStationCd, 3) = "IFC") then
	    sqlStr = "insert into [db_order].[dbo].tbl_baljumaster(baljudate,differencekey,workgroup, songjangdiv, baljutype, extSiteName, isGiftPojang, boxType, pickingStationCd, stationCdArr)"
	    sqlStr = sqlStr + " values(getdate()," + CStr(differencekey) + ",'" + workgroup + "','" + songjangdiv + "','" + baljutype + "', '" + CStr(extSiteName) + "', '" & isGiftPojang & "', '" & boxType & "', '" & pickingStationCd & "', '" & pickingStationCdArrStr & "')"
    else
	    sqlStr = "insert into [db_order].[dbo].tbl_baljumaster(baljudate,differencekey,workgroup, songjangdiv, baljutype, extSiteName, isGiftPojang, boxType, pickingStationCd)"
	    sqlStr = sqlStr + " values(getdate()," + CStr(differencekey) + ",'" + workgroup + "','" + songjangdiv + "','" + baljutype + "', '" + CStr(extSiteName) + "', '" & isGiftPojang & "', '" & boxType & "', '" & pickingStationCd & "')"
    end if


	rsget.Open sqlStr,dbget,1

	sqlStr = "select top 1 id, convert(varchar(19),baljudate,21) as baljudate from [db_order].[dbo].tbl_baljumaster order by id desc"
	rsget.Open sqlStr,dbget,1
	    iid = rsget("id")
	    obaljudate = rsget("baljudate")
	rsget.Close

	'#######출고지시 디테일###############
	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			' 정기구독 실제 발송 상품이 매칭이 안되어 있을경우 제낌.		'2019.02.13 한용민
			if instr(notmatchingstandingorder,orderserial(i)) < 1 then
				sqlStr = "insert into [db_order].[dbo].tbl_baljudetail(baljuid,orderserial,sitename,userid)"
				sqlStr = sqlStr + " values(" + CStr(iid) + ","
				sqlStr = sqlStr + " '" + orderserial(i) + "',"
				sqlStr = sqlStr + " '" + sitename(i) + "',"
				sqlStr = sqlStr + " '')"
				rsget.Open sqlStr,dbget,1
			end if
		end if
	next

    ''** [db_order].[dbo].tbl_baljudetail.baljusongjangno is NULL 인경우 업체배송으로 인식 (Logics 기존 시스템)
    ''텐바이텐 배송의 경우 송장번호를 not null 값으로 입력..

    sqlStr = "update [db_order].[dbo].tbl_baljudetail" + VbCrlf
	sqlStr = sqlStr + " set baljusongjangno=''"
	sqlStr = sqlStr + " where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " and orderserial in "
	sqlStr = sqlStr + " (select distinct bd.orderserial "
	sqlStr = sqlStr + "     from [db_order].[dbo].tbl_baljudetail bd, "
	sqlStr = sqlStr + "     [db_order].[dbo].tbl_order_detail od "
	sqlStr = sqlStr + "     where bd.baljuid=" + CStr(iid)
	sqlStr = sqlStr + "     and bd.orderserial=od.orderserial "
	sqlStr = sqlStr + "     and od.isupchebeasong='N' "
	sqlStr = sqlStr + "     and od.itemid<>0 and "
	sqlStr = sqlStr + "     od.cancelyn<>'Y' "
	sqlStr = sqlStr + "  ) "

	rsget.Open sqlStr,dbget,1

	'// 해외배송(EMS) 는 무조건 텐배
    sqlStr = " update d "
	sqlStr = sqlStr + " 	set d.baljusongjangno='' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_baljumaster m "
	sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_baljudetail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.baljuid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.id = " & CStr(iid)
	sqlStr = sqlStr + " 	and m.songjangdiv = '90' "
	sqlStr = sqlStr + " 	and d.baljusongjangno is NULL "
	''Response.write sqlStr
	rsget.Open sqlStr,dbget,1

	'// 해외배송(UPS) 는 무조건 텐배
    sqlStr = " update d "
	sqlStr = sqlStr + " 	set d.baljusongjangno='' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_baljumaster m "
	sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_baljudetail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.baljuid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.id = " & CStr(iid)
	sqlStr = sqlStr + " 	and m.songjangdiv = '92' "
	sqlStr = sqlStr + " 	and d.baljusongjangno is NULL "
	''Response.write sqlStr
	rsget.Open sqlStr,dbget,1
end if


If Err.Number = 0 Then
        errcode = "002"

	''' 플라워 주문건 CASE - 먼저 확인후 출고된경우;
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.ipkumdiv>4"
	''sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.ipkumdiv<8"  ''기 출고 완료 된것도 출고지시..
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.baljudate is NULL"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1


	'#######주문 통보 진행 ############### (출고지시일 지정)
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set ipkumdiv='5'"
	sqlStr = sqlStr + " ,baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.ipkumdiv=4"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1


	''' 업체배송 확인이 위 두 쿼리 사이에 끼이게 되면 출고지시일이 입력이 안되는 경우가 발생한다.
	''' 출고지시일 미입력된 모든 주문에 출고지시일 입력
	''' 10042631803
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.baljudate is NULL"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1

end if



If Err.Number = 0 Then
        errcode = "003"

    '#######  택배사입력 (Master) ############### (송장은 출고시 입력으로 변경 , 택배사만 입력함.)
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set songjangdiv=" + CStr(songjangdiv)
	sqlStr = sqlStr + " where orderserial in "
	sqlStr = sqlStr + " (select distinct bd.orderserial "
	sqlStr = sqlStr + "     from [db_order].[dbo].tbl_baljudetail bd, "
	sqlStr = sqlStr + "     [db_order].[dbo].tbl_order_detail od "
	sqlStr = sqlStr + "     where bd.baljuid=" + CStr(iid)
	sqlStr = sqlStr + "     and bd.orderserial=od.orderserial "
	sqlStr = sqlStr + "     and od.isupchebeasong='N' "
	sqlStr = sqlStr + "     and od.itemid<>0 and "
	sqlStr = sqlStr + "     od.cancelyn<>'Y' "
	sqlStr = sqlStr + "  ) "

	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "004"
	'###### Order Detail 텐바이텐 배송 출고지시일 저장 ############
	sqlStr = "update [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " set upcheconfirmdate= '" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " ,currstate='2'"
	sqlStr = sqlStr + " ,songjangdiv=" + CStr(songjangdiv)
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_detail.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.isupchebeasong='N'"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemid<>0"

	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "005"

	'###### Order Detail 업체 배송 업체주문통보 flag 변경 : NULL 인 경우만 ############
    sqlStr = "update [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_detail.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.isupchebeasong='Y'"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemid<>0"
	sqlStr = sqlStr + " and IsNULL([db_order].[dbo].tbl_order_detail.currstate,0)=0"

	'' 업체배송의 경우 업체주문통보 or NULL 인경우 경우 확인 가능함.
	rsget.Open sqlStr,dbget,1
end if


If Err.Number = 0 Then
        errcode = "006"

	''출고지시 아이템 목록
	sqlStr = " insert into [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " (baljuid,itemid,itemoption,rackcode,makerid," + VbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,orgsellprice,baljuno," + VbCrlf
	sqlStr = sqlStr + " ipgono,smallimage,listimage,itemrackcode)" + VbCrlf
	sqlStr = sqlStr + " select " + CStr(iid) + ",d.itemid,d.itemoption,0,'',"
	sqlStr = sqlStr + " '', '',0, sum(d.itemno)," + VbCrlf
	sqlStr = sqlStr + " 0,'','',''" + VbCrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail lo," + VbCrlf
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m," + VbCrlf
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d" + VbCrlf
	sqlStr = sqlStr + " where lo.baljuid=" + CStr(iid)  + VbCrlf
	sqlStr = sqlStr + " and lo.orderserial=d.orderserial" + VbCrlf
	sqlStr = sqlStr + " and m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " group by d.itemid,d.itemoption" + VbCrlf
	rsget.Open sqlStr,dbget,1

end if


If Err.Number = 0 Then
    errcode = "007"

	''상품랙코드지정 - 미지정된것만

'	sqlStr = " update [db_item].[dbo].tbl_item" + VbCrlf
'    sqlStr = sqlStr + " set itemrackcode=c.prtidx" + VbCrlf
'    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c" + VbCrlf
'    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.makerid=c.userid"
'    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.itemrackcode='9999'"
'    sqlStr = sqlStr + " and c.prtidx<>'9999'"
'    sqlStr = sqlStr + " and c.prtidx<>''"

	'rsget.Open sqlStr,dbget,1
end if


If Err.Number = 0 Then
    errcode = "008"

	''상품테이블관련
	sqlStr = " update [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " set makerid = T.makerid" + VbCrlf
	sqlStr = sqlStr + " ,itemname =	 T.itemname" + VbCrlf
	sqlStr = sqlStr + " ,orgsellprice = T.sellcash" + VbCrlf
	sqlStr = sqlStr + " ,smallimage = T.smallimage" + VbCrlf
	sqlStr = sqlStr + " ,listimage = T.listimage" + VbCrlf
	sqlStr = sqlStr + " ,itemrackcode = convert(varchar(4), T.itemrackcode) " + VbCrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item T" + VbCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_baljuitem.baljuid=" + CStr(iid) + VbCrlf
	sqlStr = sqlStr + " and [db_temp].[dbo].tbl_baljuitem.itemid=T.itemid" + VbCrlf
	rsget.Open sqlStr,dbget,1
    ''Response.write sqlStr
end if

If Err.Number = 0 Then
    errcode = "009"

	sqlStr = " update [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " set itemrackcode='9999'" + VbCrlf
	sqlStr = sqlStr + " where baljuid=" + CStr(iid)  + VbCrlf
	sqlStr = sqlStr + " and (itemrackcode is null or itemrackcode='')"
	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "010"

	''브랜드테이블관련
	sqlStr = " update [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " set rackcode = IsNULL(T.prtidx,0)" + VbCrlf
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c T" + VbCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_baljuitem.baljuid=" + CStr(iid) + VbCrlf
	sqlStr = sqlStr + " and [db_temp].[dbo].tbl_baljuitem.makerid=T.userid" + VbCrlf
	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "011"

	''옵션테이블관련
	sqlStr = " update [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " set itemoptionname = IsNULL(T.optionname,'')" + VbCrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option T" + VbCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_baljuitem.baljuid=" + CStr(iid) + VbCrlf
	sqlStr = sqlStr + " and [db_temp].[dbo].tbl_baljuitem.itemid=T.itemid" + VbCrlf
	sqlStr = sqlStr + " and [db_temp].[dbo].tbl_baljuitem.itemoption=T.itemoption" + VbCrlf
	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "012"

    sqlStr = " update  [db_order].[dbo].tbl_baljumaster"
    sqlStr = sqlStr + " set songjanginputed='Y'"
    sqlStr = sqlStr + " where id=" + CStr(iid)

    rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "013"
    '' 재고 업데이트
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_balju " & iid
    dbget.execute sqlStr
end if

''If Err.Number = 0 Then
''        errcode = "013"
''    ''업체 미배송/미출고지시 수량 업데이트
''    sqlStr = " exec db_partner.[dbo].sp_Ten_Partner_Summary_Upche_Mibalju_MiBeasong ''"
''
''    dbget.Execute sqlStr
''end if

If Err.Number = 0 Then
    dbget.CommitTrans

    if (Left(pickingStationCd, 3) = "IFC") then
        '// 트랜젝션 종료 후에 실행해야 한다.
        '' pickingStationCd : IFC2/IFC3/IFC4/IFC5
        '' errcode = "014"
        '' sqlStr = " exec [db_aLogistics].[dbo].[usp_LogisticsItem_Balju_to_interface] " & iid & ", '" & pickingStationCd & "', '" & pickingStationCdArrStr & "', '" & session("ssBctId") & "' "
        '' dbget_Logistics.execute sqlStr
    end if
Else
        dbget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n관리자 문의 요망 (에러코드 : " + CStr(errcode) + ")');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
End If
on error Goto 0



''추가 작업. 트랜잭션에서 뺌.. 주문내역 잠김 시간이 오래걸림 : 사은품작성시간.
''response.write "<font color=red>출고지시 생성시 오류나는 경우 서팀장께 문의요망!! - 사은품내역 </font><br>"
    '' 오래된 내역 삭제
    sqlStr = " delete from [db_temp].[dbo].tbl_baljuitem " + VbCrlf
    sqlStr = sqlStr + " where baljuid<" + CStr(iid-100)

    dbget.execute sqlStr

''리뉴얼 임시
''sqlStr = " exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_IMSI] "& iid
''dbget.execute sqlStr

' 천백만원 기프트카드 이벤트	' 2018.03.28 한용민 생성
'if (date()>="2018-04-02") and (date()<"2018-04-18") then
'	sqlStr = "exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_GiftCard_TenDLV] " & iid
'
'	'response.write sqlStr & "<br>"
'	dbget.execute sqlStr
'end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''만약 첫구매 하자고 한다면 주석 풀기 // 8/21 출고지시건으로 종료 되었음.
'if (now()>="2017-08-14") and (now()<"2017-09-02") then
'	sqlStr = " exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_FirstOrder_TenDLV] " & iid
'	dbget.execute sqlStr
'end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 사은품 배송 -> 주문시 저장.
''사은품 텐배송
''If Err.Number = 0 Then
''        errcode = "013"
''
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker " & iid & ",'N',0"
''    dbget.execute sqlStr
''end if

''    ''사은품 업체배송
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker " & iid & ",'Y',0"
''    dbget.execute sqlStr


''    야후 이벤트
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',9098,'10x10'"
''    dbget.execute sqlStr
''response.write "사은품 작성 (야후) 완료<br>"

''    ''사은품 다이어리 사은품  이벤트;;
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',8752,'10x10'"
''    dbget.execute sqlStr

    ''사은품 다이어리 사은품  이벤트 2번째 변경 Case;;
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',8842,'10x10'"
''    dbget.execute sqlStr
''response.write "사은품 작성 (다이어리 이벤트) 완료<br>"


end if

function FormatStr(n,orgData)
	dim tmp
	if (n-Len(CStr(orgData))) < 0 then
		FormatStr = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	FormatStr = tmp
end Function
%>

<script language="javascript">
alert('출고지시서가 생성 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
