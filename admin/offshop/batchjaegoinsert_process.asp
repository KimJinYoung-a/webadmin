<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx
dim iid, baljucode, brandlist
dim mode, targetid, targetname, baljuid, baljuname, reguser, regname, divcode, vatinclude, yyyymmdd, comment
dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr, sellcasharr, suplycasharr, buycasharr, designerarr
	idx = requestCheckVar(request.Form("idx"),10)
	mode = requestCheckVar(request.Form("mode"),32)
	yyyymmdd = requestCheckVar(request.Form("yyyymmdd"),10)
	baljuid = requestCheckVar(request.Form("baljuid"),32)
	targetid = requestCheckVar(request.Form("targetid"),32)
	reguser = requestCheckVar(request.Form("reguser"),32)
	divcode = requestCheckVar(request.Form("divcode"),3)
	vatinclude = requestCheckVar(request.Form("vatinclude"),1)
	comment = html2db(request.Form("comment"))
	regname = session("ssBctCname")

	itemgubunarr = request.Form("itemgubunarr")
	itemidarr = request.Form("itemidarr")
	itemoptionarr = request.Form("itemoptionarr")
	itemnoarr = request.Form("itemnoarr")
	sellcasharr = request.Form("sellcasharr")
	suplycasharr = request.Form("suplycasharr")
	buycasharr = request.Form("buycasharr")
	designerarr = request.Form("designerarr")

dim i,cnt,sqlStr

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if mode="addshopjumun" then
	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("targetid") = targetid
	rsget("targetname") = targetname
	rsget("baljuid") = baljuid
	rsget("baljuname") = baljuname
	rsget("reguser") = reguser
	rsget("regname") = regname
	rsget("divcode") = divcode
	rsget("vatinclude") = vatinclude
	rsget("scheduledate") = yyyymmdd
	rsget("statecd") = "0"
	rsget("comment") = comment

	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = "SJ" + Format00(6,Right(CStr(iid),6))

	if targetid="10x10" then
		targetname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = db2html(rsget("socname_kor"))
		end if
		rsget.close
	end if

	if baljuname="" then
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + baljuid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			baljuname = db2html(rsget("socname_kor"))
		end if
		rsget.close
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + html2db(targetname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,baljuname='" + html2db(baljuname) + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemidarr = Left(itemidarr,Len(itemidarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	buycasharr = Left(buycasharr,Len(buycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemidarr = split(itemidarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemidarr)

	for i=0 to cnt
        if (trim(itemgubunarr(i)) <> "") then
    		sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
    		sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
    		sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
    		sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
    		sqlStr = sqlStr + " values(" + CStr(iid)  + "," + vbCrlf
    		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
    		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(itemidarr(i),10) + "," + vbCrlf
    		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
    		sqlStr = sqlStr + "''," + vbCrlf
    		sqlStr = sqlStr + "''," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(buycasharr(i),20) + "," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
    		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
    		sqlStr = sqlStr + "'0')"

			'response.write sqlStr
    		rsget.Open sqlStr, dbget, 1
        end if
	next

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " ,itemoptionname=[db_shop].[dbo].tbl_shop_item.shopitemoptionname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1

    sqlStr = "update [db_shop].[dbo].tbl_shop_tempstock_master "
    sqlStr = sqlStr + " set joblinkcode = '" + CStr(baljucode) + "' "
    sqlStr = sqlStr + " where idx = " + CStr(idx) + " "
    rsget.Open sqlStr,dbget,1

end if

%>
<script type='text/javascript'>

	alert('저장 되었습니다.');
	opener.location.href=opener.document.location;
	opener.focus();
	window.close();

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
