<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 180
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode
dim yyyy,mm
dim yyyymm, yyyymmNext, idx, yyyymmPre
dim targetGbn
dim sitename, yyyymmdd, beasongPay

mode = request.Form("mode")
yyyy = request.Form("yyyy")
mm = request.Form("mm")
yyyymm = yyyy + "-" + mm
yyyymmNext = Left(CStr(DateSerial(yyyy,mm+1,1)),7)
yyyymmPre  = Left(CStr(DateSerial(yyyy,mm-1,1)),7)

idx = request.Form("idx")
targetGbn = request.Form("targetGbn")

Dim upchepro : upchepro = requestCheckvar(request("upchepro"),10)
Dim cpnidx : cpnidx = requestCheckvar(request("cpnidx"),32)
Dim differencekey : differencekey = requestCheckvar(request("differencekey"),10)
Dim makerid : makerid = requestCheckvar(request("makerid"),32)

if (request("jyyyymm")<>"") then yyyymm=request("jyyyymm")

sitename = requestCheckvar(request("sitename"),32)
yyyymmdd = requestCheckvar(request("yyyymmdd"),32)
beasongPay = requestCheckvar(request("beasongPay"),32)


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, objCmd, returnValue, retErrText

if (mode="brandcpn") then
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[usp_Ten_jungsanMake_BrandCpnSubstract]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@jyyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@upchepro", adInteger, adParamInput, , upchepro)
		.Parameters.Append .CreateParameter("@differencekey", adInteger, adParamInput, , differencekey)
		.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value
		retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	IF (returnValue<1) then
	%>
		<script language="javascript">
		alert('ERR-<%=retErrText%>');
		location.replace('<%= refer %>');
		</script>
	<%
		dbget.close():response.end
	end if

elseif (mode="addextbeasongPay") then
    '//
    sqlStr = " exec [db_order].[dbo].[usp_Ten_RegExtBeasongPayOrder] '" & sitename & "', '" & yyyymmdd & "', " & beasongPay
    dbget.Execute sqlStr

elseif (mode="brandcpnidx") then
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[usp_Ten_jungsanMake_BrandCpnSubstractByCPnIDX]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@jyyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@upchepro", adInteger, adParamInput, , upchepro)
		.Parameters.Append .CreateParameter("@differencekey", adInteger, adParamInput, , differencekey)
        .Parameters.Append .CreateParameter("@cpnIdx", adInteger, adParamInput, , cpnidx)
		.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value
		retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	IF (returnValue<1) then
	%>
		<script language="javascript">
		alert('ERR-<%=retErrText%>');
		location.replace('<%= refer %>');
		</script>
	<%
		dbget.close():response.end
	end if

elseif mode="addonebatchOF" then

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[sp_Ten_jungsanMakeByBrandOFF]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@jGubun", adVarchar, adParamInput, 2, request("jGubun"))
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@yyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@itemvatYN", adVarchar, adParamInput, 1, request("vatyn"))
		.Parameters.Append .CreateParameter("@differencekey", adInteger, adParamInput, , 0)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value

	Set objCmd = nothing

	IF (returnValue<1) then
		response.write "<script>parent.addResultLog("&request("oseq")&",'ERR:"&returnValue&"');parent.fnNextJungsanInputProc();</script>"
		dbget.close():response.end
	else

		sqlStr = "exec db_jungsan.dbo.usp_TEN_JungsanBatch_finTargetList "&returnValue&",'"&yyyymm&"','"&"OF"&"','"&request("jGubun")&"','"&request("DLVGbn")&"','"&request("vatyn")&"','"&makerid&"'"
		dbget.Execute sqlStr

		response.write "<script>parent.addResultLog("&request("oseq")&",'OK:"&returnValue&"');parent.fnNextJungsanInputProc();</script>"
		dbget.close():response.end

	end if
elseif mode="addonebatchON" then

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[sp_Ten_jungsanMakeByBrandONN]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@jGubun", adVarchar, adParamInput, 2, request("jGubun"))
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@yyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@itemvatYN", adVarchar, adParamInput, 1, request("vatyn"))
		.Parameters.Append .CreateParameter("@differencekey", adInteger, adParamInput, , 0)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value

	Set objCmd = nothing

	IF (returnValue<1) then
		response.write "<script>parent.addResultLog("&request("oseq")&",'ERR:"&returnValue&"');parent.fnNextJungsanInputProc();</script>"
		dbget.close():response.end
	else

		sqlStr = "exec db_jungsan.dbo.usp_TEN_JungsanBatch_finTargetList "&returnValue&",'"&yyyymm&"','"&"ON"&"','"&request("jGubun")&"','"&request("DLVGbn")&"','"&request("vatyn")&"','"&makerid&"'"
		dbget.Execute sqlStr

		response.write "<script>parent.addResultLog("&request("oseq")&",'OK:"&returnValue&"');parent.fnNextJungsanInputProc();</script>"
		dbget.close():response.end

	end if
elseif mode="etcChulgoJOne" then

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[sp_Ten_jungsanMakeByBrandONN_ETCCHULGO]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@mayjacctcd", adVarchar, adParamInput, 16, request("mayjacctcd"))
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@yyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@itemvatYN", adVarchar, adParamInput, 1, request("vatyn"))
		.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value
		retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	IF (returnValue<1) then
	%>
		<script language="javascript">
		alert('ERR-<%=retErrText%>');
		location.replace('<%= refer %>');
		</script>
	<%
		dbget.close():response.end
	end if
elseif mode="etcSaleMarginJOne" then

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[sp_Ten_jungsanMakeByBrandONN_MeaipSaleMarginShare]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@yyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value
		retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	IF (returnValue<1) then
	%>
		<script language="javascript">
		alert('ERR-<%=retErrText%>');
		location.replace('<%= refer %>');
		</script>
	<%
		dbget.close():response.end
	end if
elseif mode="etcBeasongPayShareOne" then

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_jungsan].[dbo].[sp_Ten_jungsanMakeByBrandONN_BeasongpayShare]"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Parameters.Append .CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
		.Parameters.Append .CreateParameter("@yyyymm", adVarchar, adParamInput, 7, yyyymm)
		.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)

		.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd.Parameters("RETURN_VALUE").Value
		retErrText  = objCmd.Parameters("@retErrText").Value
	Set objCmd = nothing

	IF (returnValue<1) then
	%>
		<script language="javascript">
		alert('ERR-<%=retErrText%>');
		location.replace('<%= refer %>');
		</script>
	<%
		dbget.close():response.end
	end if
elseif mode="makebatchtarget" then
	if (targetGbn="ON") then
		sqlStr = "exec db_jungsan.[dbo].[usp_TEN_JungsanBatch_TargetMake_ON] '"&yyyymm&"'"
		dbget.Execute sqlStr

		response.write "<script>alert('OK');parent.location.reload();</script>"
		dbget.close():response.end
	elseif (targetGbn="OF") then
		sqlStr = "exec db_jungsan.[dbo].[usp_TEN_JungsanBatch_TargetMake_OFF] '"&yyyymm&"'"
		dbget.Execute sqlStr

		response.write "<script>alert('OK');parent.location.reload();</script>"
		dbget.close():response.end
	else
		response.write "<script>alert('미지정:"&targetGbn&"')</script>"
		dbget.close():response.end
	end if

elseif mode="maeip_notax" then
	''StepI 정산 Master 2 면세 (taxtype '02', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select distinct d.imakerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='maeip'"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.divcode='001'"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and convert(varchar(7),m.executedt,20)='" + yyyymm + "'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and d.deldt is NULL"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=0"
	sqlStr = sqlStr + " )"


	rsget.Open sqlStr,dbget,1

	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,itemid,itemoption,"
	sqlStr = sqlStr + " itemname,itemoptionname,itemno,sellcash,suplycash)"
	sqlStr = sqlStr + " select jm.id, 'maeip', d.id, d.mastercode, d.itemid, d.itemoption,"
	sqlStr = sqlStr + "  d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno,d.sellcash, d.suplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm, "
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='maeip'"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.divcode='001'"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and convert(varchar(7),m.executedt,20)='" + yyyymm + "'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and d.deldt is NULL"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"  '' i.makerid -> d.imakerid로 변경
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " order by d.id desc"

	rsget.Open sqlStr,dbget,1


	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set me_cnt=T.cnt"
	sqlStr = sqlStr + " ,me_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,me_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='02'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='maeip'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"
	rsget.Open sqlStr,dbget,1



elseif mode="maeip_tax" then

	''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select distinct d.imakerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='maeip'"

	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.divcode='001'"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and convert(varchar(7),m.executedt,20)='" + yyyymm + "'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and d.deldt is NULL"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1

	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,itemid,itemoption,"
	sqlStr = sqlStr + " itemname,itemoptionname,itemno,sellcash,suplycash)"
	sqlStr = sqlStr + " select jm.id, 'maeip', d.id, d.mastercode, d.itemid, d.itemoption,"
	sqlStr = sqlStr + "  d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno,d.sellcash, d.suplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm, "
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m, "
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='maeip'"

	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.divcode='001'"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and convert(varchar(7),m.executedt,20)='" + yyyymm + "'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and d.deldt is NULL"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"   '' i.makerid -> d.imakerid 로 변경
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " order by d.id desc"

	rsget.Open sqlStr,dbget,1


	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set me_cnt=T.cnt"
	sqlStr = sqlStr + " ,me_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,me_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='01'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='maeip'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1


elseif mode="upche_notax" then
	''StepI 정산 Master 2 면세 (taxtype '02', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"

	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='Y'"
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=0"
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1

	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'upche', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='Y'"
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.makerid"  ''i.makerid -> d.makerid
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

	rsget.Open sqlStr,dbget,1

	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='02'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1

elseif mode="upche_tax" then
'response.write "관리자 문의 요망"
'dbget.close()	:	response.End

	''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid=i.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='Y'"
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " )"


dbget.Execute sqlStr
response.end



	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'upche', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"

	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d "
    sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
    sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
    sqlStr = sqlStr + "     on d.itemid=i.itemid"
    sqlStr = sqlStr + "     and i.vatinclude='Y'"			''과세
    sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_designer_jungsan_master jm"
    sqlStr = sqlStr + "     on jm.yyyymm='" + yyyymm + "'"
    sqlStr = sqlStr + "     and jm.designerid=d.makerid"
    sqlStr = sqlStr + "     and jm.taxtype='01'"
    sqlStr = sqlStr + "     and jm.differencekey=0"
    sqlStr = sqlStr + "     and jm.finishflag='0'"
    sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t "
    sqlStr = sqlStr + "     on d.idx=t.detailidx"
    sqlStr = sqlStr + "     and t.gubuncd='upche' "
    sqlStr = sqlStr + " where m.regdate>'2009-11-01'" ''datediff(m,m.regdate,'" + yyyymm + "-01')<4"
    sqlStr = sqlStr + " and m.ipkumdiv>3"
    sqlStr = sqlStr + " and m.cancelyn='N'"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.isupchebeasong='Y'"		''업체배송
    sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"


''	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
''	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master jm, "
''	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
''	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
''	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
''	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='upche'"
''	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
''	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
''	sqlStr = sqlStr + " and m.ipkumdiv>3"
''	sqlStr = sqlStr + " and m.cancelyn='N'"
''	sqlStr = sqlStr + " and d.itemid=i.itemid"
''	sqlStr = sqlStr + " and i.vatinclude='Y'"			''과세
''	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
''	sqlStr = sqlStr + " and d.itemid<>0"
''	sqlStr = sqlStr + " and d.isupchebeasong='Y'"		''업체배송
''	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
''	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
''	sqlStr = sqlStr + " and jm.designerid=d.makerid"
''	sqlStr = sqlStr + " and jm.taxtype='01'"
''	sqlStr = sqlStr + " and jm.differencekey=0"
''	sqlStr = sqlStr + " and jm.finishflag='0'"
''	sqlStr = sqlStr + " and t.detailidx is NULL"

''웹에서 타임아웃..
'response.write sqlStr
''dbget.Execute sqlStr
'response.end

	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='01'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	dbget.Execute sqlStr

elseif mode="witaksell_notax" then
	''StepI 정산 Master 2 면세 (taxtype '02', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"

	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + "     on d.itemid=i.itemid"
	sqlStr = sqlStr + "     and i.vatinclude='N'"

	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='witaksell'"

	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<2"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='N'"
	sqlStr = sqlStr + " and d.omwdiv='W'" '' -> i.mwdiv<>'M'
	sqlStr = sqlStr + " and (d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01')"
	'sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=0"
	sqlStr = sqlStr + " )"

    dbget.Execute sqlStr


'	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
'	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
'	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"

'	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
'	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
'	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
'	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='witaksell'"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_master j"
'	sqlStr = sqlStr + " on j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
'	sqlStr = sqlStr + " 	and d.makerid=j.designerid"

'	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<2"
'	sqlStr = sqlStr + " and m.ipkumdiv>3"
'	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
'	sqlStr = sqlStr + " and m.cancelyn='N'"
'	sqlStr = sqlStr + " and d.itemid=i.itemid"
'	sqlStr = sqlStr + " and i.vatinclude='N'"
'	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'	sqlStr = sqlStr + " and d.itemid<>0"
'	sqlStr = sqlStr + " and d.isupchebeasong='N'"
'	sqlStr = sqlStr + " and i.mwdiv='W'"
'	sqlStr = sqlStr + " and t.detailidx is NULL"
'	sqlStr = sqlStr + " and j.designerid is null"



	''StepII 정산 Detail
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'witaksell', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + "     on d.itemid=i.itemid"
	sqlStr = sqlStr + "     and i.vatinclude='N'"
	sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_designer_jungsan_master jm "
	sqlStr = sqlStr + "     on jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "     and jm.designerid=d.makerid"   '' i.makerid -> d.makerid

	'sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	'sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='witaksell'"

	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<2"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.omwdiv='W'"					'' -> i.mwdiv<>'M'
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') )"
	'sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"

	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	'sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

	dbget.Execute sqlStr



	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set wi_cnt=T.cnt"
	sqlStr = sqlStr + " ,wi_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,wi_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='02'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='witaksell'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	dbget.Execute sqlStr

elseif mode="witaksell_tax" then

'response.write "관리자 문의 요망..(수정필요 2008.01.02)"
'dbget.close()	:	response.End

	''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select A.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ( "
	sqlStr = sqlStr + "     select distinct d.makerid"
	sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "         Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "         on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "         Join [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + "         on d.itemid=i.itemid"
	sqlStr = sqlStr + "         and i.vatinclude='Y'"
	sqlStr = sqlStr + "         left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "         on d.idx=t.detailidx and t.gubuncd='witaksell'"

	sqlStr = sqlStr + "     where datediff(m,m.regdate,'" + yyyymm + "-01')<2"
	sqlStr = sqlStr + "     and m.ipkumdiv>3"
	sqlStr = sqlStr + "     and m.cancelyn='N'"
	sqlStr = sqlStr + "     and d.cancelyn<>'Y'"
	sqlStr = sqlStr + "     and d.itemid<>0"
	sqlStr = sqlStr + "     and d.isupchebeasong='N'"
	sqlStr = sqlStr + "     and d.omwdiv='W'"							'' -> i.mwdiv<>'M'
	sqlStr = sqlStr + "     and (d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01')"
	'sqlStr = sqlStr + "     and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + "     and t.detailidx is NULL"
	sqlStr = sqlStr + " ) A"
	sqlStr = sqlStr + " left join ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " on A.makerid=T.designerid"
    sqlStr = sqlStr + " where T.designerid is NULL"


dbget.Execute sqlStr
response.End

'dbget.close()	:	response.End
'	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
'	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
'	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"

'	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
'	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
'	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
'	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='witaksell'"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_master j"
'	sqlStr = sqlStr + " on j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
'	sqlStr = sqlStr + " 	and d.makerid=j.designerid"
'	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
'	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<2"
'	sqlStr = sqlStr + " and m.ipkumdiv>3"
'	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
'	sqlStr = sqlStr + " and m.cancelyn='N'"
'	sqlStr = sqlStr + " and d.itemid=i.itemid"
'	sqlStr = sqlStr + " and i.vatinclude='Y'"
'	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'	sqlStr = sqlStr + " and d.itemid<>0"
'	sqlStr = sqlStr + " and d.isupchebeasong='N'"
'	sqlStr = sqlStr + " and i.mwdiv='W'"
'	sqlStr = sqlStr + " and t.detailidx is NULL"
'	sqlStr = sqlStr + " and j.designerid is null"



	''StepII 정산 Detail
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'witaksell', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + "     on d.itemid=i.itemid"
	sqlStr = sqlStr + "     and i.vatinclude='Y'"
	sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_designer_jungsan_master jm "
	sqlStr = sqlStr + "     on jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "     and jm.designerid=d.makerid"   'i.makerid->d.makerid
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.idx=t.detailidx and t.gubuncd='witaksell'"

	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<2"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	'sqlStr = sqlStr + " and ((m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01') )"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and d.isupchebeasong='N'"
	sqlStr = sqlStr + " and d.omwdiv='W'"                                 '' -> i.mwdiv<>'M'
	sqlStr = sqlStr + " and (d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') "
	''sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	''sqlStr = sqlStr + " order by d.orderserial desc"


'response.write sqlStr
	'dbget.Execute sqlStr
'response.end

	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set wi_cnt=T.cnt"
	sqlStr = sqlStr + " ,wi_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,wi_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='01'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='witaksell'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	dbget.Execute sqlStr

elseif mode="maeipchulgo" then
	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if

	''StepI 정산 Master -1 과세.. (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
	sqlStr = sqlStr + " select distinct d.imakerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='maeipchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " )"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"
	sqlStr = sqlStr + " select jm.id , 'maeipchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash"

	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='maeipchulgo'"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepI 정산 Master -2 면세.. (taxtype '02', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
	sqlStr = sqlStr + " select distinct d.imakerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='maeipchulgo'"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=0"
	sqlStr = sqlStr + " )"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"
	sqlStr = sqlStr + " select jm.id , 'maeipchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash"

	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='maeipchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1

elseif mode="witakchulgo" then
	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if


    '// ================================
    '// 온라인 상품
    '// ================================

	''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,designerid,title,finishflag,groupid,jgubun,itemvatYn"
	sqlStr = sqlStr + " ,ub_cnt,ub_totalsellcash,ub_totalsuplycash"
	sqlStr = sqlStr + " ,me_cnt,me_totalsellcash,me_totalsuplycash"
	sqlStr = sqlStr + " ,wi_cnt,wi_totalsellcash,wi_totalsuplycash"
	sqlStr = sqlStr + " ,sh_cnt,sh_totalsellcash,sh_totalsuplycash"
	sqlStr = sqlStr + " ,et_cnt,et_totalsellcash,et_totalsuplycash"
	sqlStr = sqlStr + " ,wi_totalreducedprice,ub_totalreducedprice,et_totalreducedprice"
	sqlStr = sqlStr + " ,dlv_totalsellcash,dlv_totalreducedprice,dlv_totalsuplycash"
	sqlStr = sqlStr + " ,totalcommission)"

	sqlStr = sqlStr + " select distinct '" + yyyymm + "',1,'01',d.imakerid,'" + yyyy + "년 " + mm + "월 정산',0,p.groupid,'MM','Y'"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='witakchulgo'"
    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p on d.imakerid=p.id"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=1 and j.jgubun='MM'"
	sqlStr = sqlStr + " )"


	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode ,reqname"
	sqlStr = sqlStr + " ,itemid,itemoption,itemname,itemoptionname"
	sqlStr = sqlStr + " ,itemno,sellcash,suplycash,sitename,reducedprice"
	sqlStr = sqlStr + " ,commission,iszerotax,paymethod,beasongdate,vatyn,CpnNotAppliedPrice)"

	sqlStr = sqlStr + " select jm.id , 'witakchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash, m.socid,d.sellcash,"
    sqlStr = sqlStr + " 0,NULL,NULL,convert(Varchar(10),m.executedt,21),'Y',d.sellcash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.differencekey=1" ''--1로 수정
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and jm.jgubun='MM'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1

    ''면세
    sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,designerid,title,finishflag,groupid,jgubun,itemvatYn"
	sqlStr = sqlStr + " ,ub_cnt,ub_totalsellcash,ub_totalsuplycash"
	sqlStr = sqlStr + " ,me_cnt,me_totalsellcash,me_totalsuplycash"
	sqlStr = sqlStr + " ,wi_cnt,wi_totalsellcash,wi_totalsuplycash"
	sqlStr = sqlStr + " ,sh_cnt,sh_totalsellcash,sh_totalsuplycash"
	sqlStr = sqlStr + " ,et_cnt,et_totalsellcash,et_totalsuplycash"
	sqlStr = sqlStr + " ,wi_totalreducedprice,ub_totalreducedprice,et_totalreducedprice"
	sqlStr = sqlStr + " ,dlv_totalsellcash,dlv_totalreducedprice,dlv_totalsuplycash"
	sqlStr = sqlStr + " ,totalcommission)"

	sqlStr = sqlStr + " select distinct '" + yyyymm + "',1,'02',d.imakerid,'" + yyyy + "년 " + mm + "월 정산',0,p.groupid,'MM','N'"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='witakchulgo'"
    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p on d.imakerid=p.id"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=1 and j.jgubun='MM'"
	sqlStr = sqlStr + " )"


	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode ,reqname"
	sqlStr = sqlStr + " ,itemid,itemoption,itemname,itemoptionname"
	sqlStr = sqlStr + " ,itemno,sellcash,suplycash,sitename,reducedprice"
	sqlStr = sqlStr + " ,commission,iszerotax,paymethod,beasongdate,vatyn,CpnNotAppliedPrice)"

	sqlStr = sqlStr + " select jm.id , 'witakchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash, m.socid,d.sellcash,"
    sqlStr = sqlStr + " 0,NULL,NULL,convert(Varchar(10),m.executedt,21),'N',d.sellcash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
	sqlStr = sqlStr + " and i.itemid=d.itemid"
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.differencekey=1"
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and jm.jgubun='MM'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


    '// ================================
    '// 오프라인 상품
    '// ================================

	''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,designerid,title,finishflag,groupid,jgubun,itemvatYn"
	sqlStr = sqlStr + " ,ub_cnt,ub_totalsellcash,ub_totalsuplycash"
	sqlStr = sqlStr + " ,me_cnt,me_totalsellcash,me_totalsuplycash"
	sqlStr = sqlStr + " ,wi_cnt,wi_totalsellcash,wi_totalsuplycash"
	sqlStr = sqlStr + " ,sh_cnt,sh_totalsellcash,sh_totalsuplycash"
	sqlStr = sqlStr + " ,et_cnt,et_totalsellcash,et_totalsuplycash"
	sqlStr = sqlStr + " ,wi_totalreducedprice,ub_totalreducedprice,et_totalreducedprice"
	sqlStr = sqlStr + " ,dlv_totalsellcash,dlv_totalreducedprice,dlv_totalsuplycash"
	sqlStr = sqlStr + " ,totalcommission)"

	sqlStr = sqlStr + " select distinct '" + yyyymm + "',1,'01',d.imakerid,'" + yyyy + "년 " + mm + "월 정산',0,p.groupid,'MM','Y'"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " from [db_shop].[dbo].[tbl_shop_item] i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='witakchulgo'"
    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p on d.imakerid=p.id"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
    sqlStr = sqlStr + " and d.iitemgubun<> '10' "     '''오프라인 상품
	sqlStr = sqlStr + " and d.iitemgubun=i.itemgubun "
	sqlStr = sqlStr + " and i.shopitemid=d.itemid "
	sqlStr = sqlStr + " and i.itemoption=d.itemoption "
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=1 and j.jgubun='MM'"
	sqlStr = sqlStr + " )"


	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode ,reqname"
	sqlStr = sqlStr + " ,itemid,itemoption,itemname,itemoptionname"
	sqlStr = sqlStr + " ,itemno,sellcash,suplycash,sitename,reducedprice"
	sqlStr = sqlStr + " ,commission,iszerotax,paymethod,beasongdate,vatyn,CpnNotAppliedPrice)"

	sqlStr = sqlStr + " select jm.id , 'witakchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash, m.socid,d.sellcash,"
    sqlStr = sqlStr + " 0,NULL,NULL,convert(Varchar(10),m.executedt,21),'Y',d.sellcash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_shop].[dbo].[tbl_shop_item] i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
    sqlStr = sqlStr + " and d.iitemgubun<> '10' "     '''오프라인 상품
	sqlStr = sqlStr + " and d.iitemgubun=i.itemgubun "
	sqlStr = sqlStr + " and i.shopitemid=d.itemid "
	sqlStr = sqlStr + " and i.itemoption=d.itemoption "
	sqlStr = sqlStr + " and i.vatinclude='Y'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.differencekey=1" ''--1로 수정
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and jm.jgubun='MM'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1

    ''면세
    sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,designerid,title,finishflag,groupid,jgubun,itemvatYn"
	sqlStr = sqlStr + " ,ub_cnt,ub_totalsellcash,ub_totalsuplycash"
	sqlStr = sqlStr + " ,me_cnt,me_totalsellcash,me_totalsuplycash"
	sqlStr = sqlStr + " ,wi_cnt,wi_totalsellcash,wi_totalsuplycash"
	sqlStr = sqlStr + " ,sh_cnt,sh_totalsellcash,sh_totalsuplycash"
	sqlStr = sqlStr + " ,et_cnt,et_totalsellcash,et_totalsuplycash"
	sqlStr = sqlStr + " ,wi_totalreducedprice,ub_totalreducedprice,et_totalreducedprice"
	sqlStr = sqlStr + " ,dlv_totalsellcash,dlv_totalreducedprice,dlv_totalsuplycash"
	sqlStr = sqlStr + " ,totalcommission)"

	sqlStr = sqlStr + " select distinct '" + yyyymm + "',1,'02',d.imakerid,'" + yyyy + "년 " + mm + "월 정산',0,p.groupid,'MM','N'"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0,0,0"
	sqlStr = sqlStr + " ,0"
	sqlStr = sqlStr + " from [db_shop].[dbo].[tbl_shop_item] i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx  and t.gubuncd='witakchulgo'"
    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p on d.imakerid=p.id"
	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
    sqlStr = sqlStr + " and d.iitemgubun<> '10' "     '''오프라인 상품
	sqlStr = sqlStr + " and d.iitemgubun=i.itemgubun "
	sqlStr = sqlStr + " and i.shopitemid=d.itemid "
	sqlStr = sqlStr + " and i.itemoption=d.itemoption "
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.imakerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=1 and j.jgubun='MM'"
	sqlStr = sqlStr + " )"


	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode ,reqname"
	sqlStr = sqlStr + " ,itemid,itemoption,itemname,itemoptionname"
	sqlStr = sqlStr + " ,itemno,sellcash,suplycash,sitename,reducedprice"
	sqlStr = sqlStr + " ,commission,iszerotax,paymethod,beasongdate,vatyn,CpnNotAppliedPrice)"

	sqlStr = sqlStr + " select jm.id , 'witakchulgo', d.id, d.mastercode,m.socid,"
	sqlStr = sqlStr + " d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
	sqlStr = sqlStr + " d.itemno*-1,"
	sqlStr = sqlStr + " d.sellcash, d.buycash, m.socid,d.sellcash,"
    sqlStr = sqlStr + " 0,NULL,NULL,convert(Varchar(10),m.executedt,21),'N',d.sellcash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [db_shop].[dbo].[tbl_shop_item] i,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"

	sqlStr = sqlStr + " where d.id in (" + idx + ")"
	sqlStr = sqlStr + " and m.code=d.mastercode"
    sqlStr = sqlStr + " and d.iitemgubun<> '10' "     '''오프라인 상품
	sqlStr = sqlStr + " and d.iitemgubun=i.itemgubun "
	sqlStr = sqlStr + " and i.shopitemid=d.itemid "
	sqlStr = sqlStr + " and i.itemoption=d.itemoption "
	sqlStr = sqlStr + " and i.vatinclude='N'"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
	sqlStr = sqlStr + " and jm.differencekey=1"
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and jm.jgubun='MM'"
	sqlStr = sqlStr + " and t.detailidx is NULL"

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1


'	''StepI 정산 Master 1 비과세 (taxtype '02', differencekey 0)
'	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
'	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
'	sqlStr = sqlStr + " select distinct d.imakerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"
'	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
'	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
'	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
'	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"
'
'	sqlStr = sqlStr + " where d.id in (" + idx + ")"
'	sqlStr = sqlStr + " and m.code=d.mastercode"
'	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
'	sqlStr = sqlStr + " and i.itemid=d.itemid"
'	sqlStr = sqlStr + " and i.vatinclude='N'"
'	sqlStr = sqlStr + " and t.detailidx is NULL"
'	sqlStr = sqlStr + " and d.imakerid not in ("
'	sqlStr = sqlStr + " 	select designerid"
'	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
'	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='02' and differencekey=0"
'	sqlStr = sqlStr + " )"
'
'	'response.write sqlStr
'	rsget.Open sqlStr,dbget,1
'
'
'	''StepII 정산 Detail Insert
'
'	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
'	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,reqname,"
'	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
'	sqlStr = sqlStr + " sellcash,suplycash)"
'	sqlStr = sqlStr + " select jm.id , 'witakchulgo', d.id, d.mastercode,m.socid,"
'	sqlStr = sqlStr + " d.itemid, d.itemoption,"
'	sqlStr = sqlStr + " d.iitemname, d.iitemoptionname as itemoptionname,"
'	sqlStr = sqlStr + " d.itemno*-1,"
'	sqlStr = sqlStr + " d.sellcash, d.buycash"
'
'	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
'	sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
'	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
'	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
'	sqlStr = sqlStr + " left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
'	sqlStr = sqlStr + " on d.id=t.detailidx and t.gubuncd='witakchulgo'"
'
'	sqlStr = sqlStr + " where d.id in (" + idx + ")"
'	sqlStr = sqlStr + " and m.code=d.mastercode"
'	sqlStr = sqlStr + " and d.iitemgubun='10'"        '''온라인 상품코드만..
'	sqlStr = sqlStr + " and i.itemid=d.itemid"
'	sqlStr = sqlStr + " and i.vatinclude='N'"
'	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
'	sqlStr = sqlStr + " and jm.designerid=d.imakerid"
'	sqlStr = sqlStr + " and jm.differencekey=0"
'	sqlStr = sqlStr + " and jm.taxtype='02'"
'	sqlStr = sqlStr + " and jm.finishflag='0'"
'	sqlStr = sqlStr + " and t.detailidx is NULL"
'
'	'response.write sqlStr
'	rsget.Open sqlStr,dbget,1



'	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set et_cnt=T.cnt"
	sqlStr = sqlStr + " ,et_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,et_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select m.id, count(d.id) as cnt,"
	sqlStr = sqlStr + " 	sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " 	where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " 	and m.id=d.masteridx"
	sqlStr = sqlStr + " 	and d.gubuncd='witakchulgo'"
	sqlStr = sqlStr + " 	and m.jgubun='MM'"
	sqlStr = sqlStr + " 	group by m.id"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.jgubun='MM'"

	rsget.Open sqlStr,dbget,1
elseif mode="lectureipkumdivproc" then
	sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
	sqlStr = sqlStr + " set beadaldate=T.findate" + VbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='7'" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + " 	select distinct  d.mastercode, dateadd(d,-1,dateAdd(m,1,m.yyyymm+'-01')) as findate" + VbCrlf
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master m," + VbCrlf
	sqlStr = sqlStr + " 	[db_jungsan].[dbo].tbl_designer_jungsan_detail d" + VbCrlf
	sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item i" + VbCrlf
	sqlStr = sqlStr + " 	on d.itemid=i.itemid" + VbCrlf
	sqlStr = sqlStr + " 	where m.id=d.masteridx" + VbCrlf
	sqlStr = sqlStr + " 	and m.yyyymm='" + yyyymm + "'" + VbCrlf
	sqlStr = sqlStr + " 	and d.gubuncd='upche'" + VbCrlf
	sqlStr = sqlStr + " 	and i.itemdiv='90'" + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial=T.mastercode" + VbCrlf
	sqlStr = sqlStr + " and  cancelyn='N'" + VbCrlf
	sqlStr = sqlStr + " and ipkumdiv<7" + VbCrlf
	sqlStr = sqlStr + " and ipkumdiv>4" + VbCrlf

	rsget.Open sqlStr,dbget,1

elseif mode="finishflag1" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
	sqlStr = sqlStr + " set finishflag='1'" + VbCrlf
	sqlStr = sqlStr + " where yyyymm='" + yyyymm + "'" + VbCrlf
	sqlStr = sqlStr + " and finishflag='0'" + VbCrlf
    sqlStr = sqlStr + " and ub_totalsuplycash+me_totalsuplycash+wi_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash<>0"

	dbget.Execute sqlstr

elseif (mode="finishflagoff1") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='1'" + VbCrlf
    sqlstr = sqlstr + " where yyyymm='" + CStr(yyyymm) + "'" + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"  + VbCrlf
    sqlstr = sqlstr + " and tot_jungsanprice<>0"  + VbCrlf

    dbget.Execute sqlstr

elseif mode="upchedeliverPay" then
    rw "사용안함"
    response.end
'response.write "배송일로 수정요망."
'response.write "상품의 최종배송일 기준으로.."
'response.end
'dbget.close()	:	response.End
''lovepiary 면세.. 관련. 검토 :: jm.taxtype='02'로 함더 돌려야함.

    ''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     on d.makerid=c.userid"
	sqlStr = sqlStr + "     and c.defaultDeliverPay>0"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "         select distinct d.mastercode"
    sqlStr = sqlStr + "         from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
    sqlStr = sqlStr + "     	    Join [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
    sqlStr = sqlStr + "     	    on m.id=d.masteridx"
    sqlStr = sqlStr + "     	    Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     	    on m.designerid=c.userid"
	sqlStr = sqlStr + "     	    and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "         and m.yyyymm='"+yyyymm+"'"
    sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "     ) T2 "
	sqlStr = sqlStr + "     on d.orderserial=T2.mastercode"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.idx=t.detailidx and t.gubuncd='upche'"
	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=0"
	sqlStr = sqlStr + " and d.buycash>0"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01')"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
	''sqlStr = sqlStr + " and d.isupchebeasong='Y'"
	''sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " )"


	dbget.Execute sqlStr
''response.write sqlStr
''response.end

	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"
	sqlStr = sqlStr + " select jm.id, 'upche', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
''sqlStr = sqlStr + "     and d.makerid='mungunara'"
	sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_designer_jungsan_master jm "
	sqlStr = sqlStr + "     on jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "     and jm.designerid=d.makerid"
	sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     on d.makerid=c.userid"
	sqlStr = sqlStr + "     and c.defaultDeliverPay>0"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "         select distinct d.mastercode"
    sqlStr = sqlStr + "         from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
    sqlStr = sqlStr + "     	    Join [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
    sqlStr = sqlStr + "     	    on m.id=d.masteridx"
    sqlStr = sqlStr + "     	    Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     	    on m.designerid=c.userid"
	sqlStr = sqlStr + "     	    and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "         and m.yyyymm='"+yyyymm+"'"
    sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "     ) T2 "
	sqlStr = sqlStr + "     on d.orderserial=T2.mastercode"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "     select m.designerid, d.mastercode "
    sqlStr = sqlStr + "         from db_jungsan.dbo.tbl_designer_jungsan_master m"
	sqlStr = sqlStr + "         Join db_jungsan.dbo.tbl_designer_jungsan_detail d"
	sqlStr = sqlStr + "         on m.id=d.masteridx"
	sqlStr = sqlStr + "         and m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "         Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "         on m.designerid=c.userid"
	sqlStr = sqlStr + "         and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "     group by m.designerid, d.mastercode "
    sqlStr = sqlStr + "     ) T3 on d.orderserial=T3.mastercode"
    sqlStr = sqlStr + "     and d.makerid=T3.designerid"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.idx=t.detailidx and t.gubuncd='upche'"
	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=0"
	sqlStr = sqlStr + " and d.buycash>0"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01')"  '''이부분확인
	''sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
	''sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

''response.write sqlStr
'response.end
	dbget.Execute sqlStr


	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='01'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	dbget.Execute sqlStr
elseif mode="upchedeliverPay_notax" then
     rw "사용안함"
    response.end
'response.write "배송일로 수정요망."
'response.write "상품의 최종배송일 기준으로.."
'response.end
'dbget.close()	:	response.End
''lovepiary 면세.. 관련. 검토 :: jm.taxtype='02'로 함더 돌려야함.

    ''StepI 정산 Master 1 과세 (taxtype '01', differencekey 0)
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"

	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','02',0"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     on d.makerid=c.userid"
	sqlStr = sqlStr + "     and c.defaultDeliverPay>0"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "         select distinct d.mastercode"
    sqlStr = sqlStr + "         from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
    sqlStr = sqlStr + "     	    Join [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
    sqlStr = sqlStr + "     	    on m.id=d.masteridx"
    sqlStr = sqlStr + "     	    Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     	    on m.designerid=c.userid"
	sqlStr = sqlStr + "     	    and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "         and m.yyyymm='"+yyyymm+"'"
    sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "     ) T2 "
	sqlStr = sqlStr + "     on d.orderserial=T2.mastercode"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.idx=t.detailidx and t.gubuncd='upche'"
	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=0"
	sqlStr = sqlStr + " and d.buycash>0"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01')"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
	''sqlStr = sqlStr + " and d.isupchebeasong='Y'"
	''sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "'  and differencekey=0"
	sqlStr = sqlStr + " )"


	dbget.Execute sqlStr
response.write sqlStr
''response.end

	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'upche', d.idx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
	sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_designer_jungsan_master jm "
	sqlStr = sqlStr + "     on jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "     and jm.designerid=d.makerid"
	sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     on d.makerid=c.userid"
	sqlStr = sqlStr + "     and c.defaultDeliverPay>0"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "         select distinct d.mastercode"
    sqlStr = sqlStr + "         from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
    sqlStr = sqlStr + "     	    Join [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
    sqlStr = sqlStr + "     	    on m.id=d.masteridx"
    sqlStr = sqlStr + "     	    Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "     	    on m.designerid=c.userid"
	sqlStr = sqlStr + "     	    and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "         and m.yyyymm='"+yyyymm+"'"
    sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "     ) T2 "
	sqlStr = sqlStr + "     on d.orderserial=T2.mastercode"
	sqlStr = sqlStr + "     Join ("
	sqlStr = sqlStr + "     select m.designerid, d.mastercode "
    sqlStr = sqlStr + "         from db_jungsan.dbo.tbl_designer_jungsan_master m"
	sqlStr = sqlStr + "         Join db_jungsan.dbo.tbl_designer_jungsan_detail d"
	sqlStr = sqlStr + "         on m.id=d.masteridx"
	sqlStr = sqlStr + "         and m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "         and d.gubuncd='upche'"
	sqlStr = sqlStr + "         Join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + "         on m.designerid=c.userid"
	sqlStr = sqlStr + "         and c.defaultDeliverPay>0"
    sqlStr = sqlStr + "     group by m.designerid, d.mastercode "
    sqlStr = sqlStr + "     ) T3 on d.orderserial=T3.mastercode"
    sqlStr = sqlStr + "     and d.makerid=T3.designerid"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.idx=t.detailidx and t.gubuncd='upche'"
	sqlStr = sqlStr + " where datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid=0"
	sqlStr = sqlStr + " and d.buycash>0"
	sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01')"
	''sqlStr = sqlStr + " and (m.beadaldate>='" + yyyymm + "-01' and m.beadaldate<'" + yyyymmNext + "-01')"
	''sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') or (m.jumundiv='9' and m.regdate>='" + yyyymm + "-01' and m.regdate<'" + yyyymmNext + "-01' ))"
	sqlStr = sqlStr + " and jm.taxtype='02'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

'response.write sqlStr
'response.end
	dbget.Execute sqlStr


	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='02'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	dbget.Execute sqlStr
elseif mode="lectureBatch" then

    sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey,targetGbn)"
	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','03',0,'AC'"

	sqlStr = sqlStr + " from [110.93.128.73].[db_academy].[dbo].tbl_lec_item i,"
	sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_academy_order_master m,"
	sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_academy_order_detail d"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.detailidx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.sitename='academy'"
	sqlStr = sqlStr + " and d.itemid=i.idx"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and i.lec_date='" + yyyymm + "'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='03' and differencekey=0"
	sqlStr = sqlStr + " )"

	dbget.Execute sqlStr


	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'upche', d.detailidx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_lec_item i,"
	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_academy_order_master m,"
	sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_academy_order_detail d"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.orderserial=t.mastercode and d.detailidx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.sitename='academy'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and i.lec_date='" + yyyymm + "'"
	sqlStr = sqlStr + " and d.itemid=i.idx"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.makerid"  ''i.makerid -> d.makerid
	sqlStr = sqlStr + " and jm.taxtype='03'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

	rsget.Open sqlStr,dbget,1

	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='03'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1

elseif mode="DIYBatch" then

    sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey,targetGbn)"
	sqlStr = sqlStr + " select distinct d.makerid,'" + yyyymm + "','" + yyyy + "년 " + mm + "월 정산','01',0,'AC'"

	''sqlStr = sqlStr + " from [110.93.128.73].[db_academy].[dbo].tbl_lec_item i,"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m,"
	sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_detail d"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.detailidx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.sitename='diyitem'"
	''sqlStr = sqlStr + " and d.itemid=i.idx"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') )"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " and d.makerid not in ("
	sqlStr = sqlStr + " 	select designerid"
	sqlStr = sqlStr + " 	from [db_jungsan].[dbo].tbl_designer_jungsan_master j"
	sqlStr = sqlStr + " 	where  j.yyyymm='" + yyyymm + "' and taxtype='01' and differencekey=0"
	sqlStr = sqlStr + " )"

	dbget.Execute sqlStr

    ''--- 배송비0 코드 배송일 설정..

    sqlStr = " update D "
    sqlStr = sqlStr + " set beasongdate=T.Mbeasongdate"
    sqlStr = sqlStr + " from [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_detail D"
    sqlStr = sqlStr + " 	Join ("
    sqlStr = sqlStr + " 	select  d.orderserial, d.makerid, max(beasongdate) as Mbeasongdate "
    sqlStr = sqlStr + " 	from [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m"
    sqlStr = sqlStr + " 		Join [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_detail d"
    sqlStr = sqlStr + " 		on m.orderserial=d.orderserial"
    sqlStr = sqlStr + " 	where m.sitename='diyitem'"
    sqlStr = sqlStr + "     and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
    sqlStr = sqlStr + " 	and m.cancelyn='N'"
    sqlStr = sqlStr + " 	and m.ipkumdiv>3"
    sqlStr = sqlStr + " 	and d.cancelyn='N'"
    sqlStr = sqlStr + "     and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') )"
    sqlStr = sqlStr + " 	group by d.orderserial, d.makerid"
    sqlStr = sqlStr + " 	) T"
    sqlStr = sqlStr + "  	on D.orderserial=T.orderserial"
    sqlStr = sqlStr + " 	and D.makerid=T.makerid"
    sqlStr = sqlStr + " 	and D.itemid=0"
    sqlStr = sqlStr + " where D.beasongdate is NULL"

    rsget.Open sqlStr,dbget,1


	''StepII 정산 Detail Insert
	sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
	sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
	sqlStr = sqlStr + " sellcash,suplycash)"

	sqlStr = sqlStr + " select jm.id, 'upche', d.detailidx, d.orderserial, m.buyname, m.reqname, d.itemid, d.itemoption,"
	sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
	sqlStr = sqlStr + " d.itemcost, d.buycash"
	sqlStr = sqlStr + " from "
	''sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_lec_item i,"
	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master jm,"
	sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m,"
	sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_detail d"
	sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail t"
	sqlStr = sqlStr + "     on d.orderserial=t.mastercode and d.detailidx=t.detailidx and t.gubuncd='upche'"

	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	sqlStr = sqlStr + " and datediff(m,m.regdate,'" + yyyymm + "-01')<4"
	sqlStr = sqlStr + " and m.ipkumdiv>3"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.sitename='diyitem'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	'''sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and ((d.beasongdate>='" + yyyymm + "-01' and d.beasongdate<'" + yyyymmNext + "-01') )"
	'''sqlStr = sqlStr + " and i.lec_date='" + yyyymm + "'"
	'''sqlStr = sqlStr + " and d.itemid=i.idx"
	sqlStr = sqlStr + " and jm.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and jm.designerid=d.makerid"  ''i.makerid -> d.makerid
	sqlStr = sqlStr + " and jm.taxtype='01'"
	sqlStr = sqlStr + " and jm.differencekey=0"
	sqlStr = sqlStr + " and jm.finishflag='0'"
	sqlStr = sqlStr + " and t.detailidx is NULL"
	sqlStr = sqlStr + " order by d.orderserial desc"

	rsget.Open sqlStr,dbget,1

	''StepIII 정산 Master Summary
	sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ub_cnt=T.cnt"
	sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select m.id, count(d.id) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d, [db_jungsan].[dbo].tbl_designer_jungsan_master m"
	sqlStr = sqlStr + " where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and m.taxtype='01'"
	sqlStr = sqlStr + " and m.differencekey=0"
	sqlStr = sqlStr + " and m.id=d.masteridx"
	sqlStr = sqlStr + " and d.gubuncd='upche'"
	sqlStr = sqlStr + " group by m.id) as T"
	sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1

end if

''정산작성시 Groupid 추가  2007-02-01
    sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master"
    sqlStr = sqlStr + " set groupid=p.groupid"
    sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
    sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.yyyymm='" + yyyymm + "'"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
    sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.groupid is null"

    rsget.Open sqlStr,dbget,1
%>


<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
