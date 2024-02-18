<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���
' History : �̻� ����
'			2019.05.23 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim masterid
dim code, mode, scheduledt, executedt, chargeid, chargename, storeid, divcode, vatcode, comment
dim itemgubunarr, itemidarr, itemoptionarr, itemnamearr, itemoptionnamearr, sellcasharr, buycasharr, suplycasharr, itemnoarr, designerarr, mwdivarr
dim itemgubun, itemid, itemoption, itemname, itemoptionname, sellcash, buycash, suplycash, itemno, designer, mwdiv
dim statecd
dim menupos
Dim finishid, finishname
dim shopid, alinkcode, currencyUnit, currencyUnit_Pos, priceChange, masteridx, sitename

masterid = request("masterid")
code = request("code")
mode = request("mode")
scheduledt = request("scheduledt")
executedt = request("executedt")
chargeid = session("ssBctid")
chargename = session("ssBctCname")
storeid = request("storeid")
socname = request("socname")
divcode = request("divcode")
vatcode = request("vatcode")
comment = html2db(request("comment"))
statecd = request("statecd")

itemgubunarr = request("itemgubunarr")
itemidarr = request("itemidarr")
itemoptionarr = request("itemoptionarr")
itemnamearr = html2db(request("itemnamearr"))
itemoptionnamearr = html2db(request("itemoptionnamearr"))
sellcasharr = request("sellcasharr")
buycasharr = request("buycasharr")
suplycasharr = request("suplycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
mwdivarr = request("mwdivarr")
menupos =  request("menupos")

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnamearr = split(itemnamearr, "|")
itemoptionnamearr = split(itemoptionnamearr, "|")
sellcasharr = split(sellcasharr, "|")
buycasharr = split(buycasharr, "|")
suplycasharr = split(suplycasharr, "|")
itemnoarr = split(itemnoarr, "|")
designerarr = split(designerarr, "|")
mwdivarr = split(mwdivarr, "|")

finishid = session("ssBctid")
finishname = html2db(session("ssBctCname"))

dim i,cnt,sqlStr, chk, didx
dim AssignedRows

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim socname, iid

dim STOCKBASEDATE, yyyymmdd
dim tmpitemid, tmpitemgubun, tmpitemoption, tmpitemno, isdeleted, ipchulflag
dim puserdiv
isdeleted = false

response.write "mode=" + mode

if mode="write" then '�ֹ�����
	puserdiv=""
	sqlStr = " select top 1 isNULL(userdiv,'') as puserdiv" & vbcrlf	' ���ó üũ�߰�	' 2019.05.23 �ѿ�� �߰�
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner" & vbcrlf
	sqlStr = sqlStr & " where id='" + storeid + "'" & vbcrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		puserdiv = rsget("puserdiv")
	else
		response.write "<script type='text/javascript'>alert('�ش� �Ǵ� ���ó�� �����ϴ�.');</script>"
		response.write "<script type='text/javascript'>location.replace('"& refer &"')</script>"
		rsget.close : dbget.close() : response.End
	end if
	rsget.close

	if masterid = "" or masterid ="0" then   ' �ӽ����� ���� ���

			'��ü�� �˻�
			sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
			sqlStr = sqlStr + " where userid='" + storeid + "'"
			rsget.Open sqlStr, dbget, 1
			if Not rsget.Eof then
				socname = rsget("socname_kor")
			end if
			rsget.close

				sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
				rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("code") = ""
				rsget("socid") = storeid
				rsget("socname") = socname
				rsget("chargeid") = chargeid
				rsget("chargename") = chargename
				rsget("divcode") = divcode
				rsget("vatcode") = "008"
				rsget("comment") = comment
				rsget("scheduledt") = scheduledt
				'rsget("executedt") = executedt
				rsget("statecd") = 1

				if (Left(storeid,10)="streetshop") or (Left(storeid,9)="wholesale") or (puserdiv="501") or (puserdiv="502") or (puserdiv="503") then
					rsget("ipchulflag") = "S"
				else
					rsget("ipchulflag") = "E"
				end if


				rsget.update
				iid = rsget("id")
				rsget.close
				code = "SO" + Format00(6,Right(CStr(iid),6))

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
			sqlStr = sqlStr + " set code='" + code + "'" + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(iid)
			rsget.Open sqlStr,dbget,1

			sqlStr = " delete from [db_storage].[dbo].tbl_acount_storage_detail where mastercode= '" + CStr(code) + "'"
			dbget.execute sqlStr
	ELSE
			iid = masterid
	END IF

			'''2.�¶��� �԰� ������ �Է�
			for i=0 to UBound(itemgubunarr) - 1
				if (trim(itemgubunarr(i)) <> "") then
					itemgubun = trim(itemgubunarr(i))
					itemid = trim(itemidarr(i))
					itemoption = trim(itemoptionarr(i))
					itemname = trim(itemnamearr(i))
					itemoptionname = trim(itemoptionnamearr(i))
					sellcash = trim(sellcasharr(i))
					suplycash = trim(suplycasharr(i))
					buycash = trim(buycasharr(i))

					'���� ������ �پ��� ���̴�.
					itemno = -1 * CInt(trim(itemnoarr(i)))

					designer = trim(designerarr(i))
					mwdiv = trim(mwdivarr(i))
					itemname = ""
					itemoptionname = ""

					sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
					sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
					sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
					sqlStr = sqlStr + " values('" + code + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
					rsget.Open sqlStr,dbget,1
				end if
			next

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,socid='"+storeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,socname='"+socname+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargeid='"+chargeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargename='"+chargename+"'" + VBCrlf
	sqlStr = sqlStr + " ,divcode='"+divcode+"'" + VBCrlf
	sqlStr = sqlStr + " ,vatcode='008'" + VBCrlf
	sqlStr = sqlStr + " ,comment='"+comment+"'" + VBCrlf
	sqlStr = sqlStr + " ,scheduledt='"+scheduledt+"'" + VBCrlf
	sqlStr = sqlStr + " ,statecd=1" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(code) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('��� ������ �ԷµǾ����ϴ�. - ��� Ȯ���ϼž� ���ó���˴ϴ�.');</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=" + storeid + "&statecd=1&menupos="+menupos+"')</script>"
	dbget.close()	:	response.End

elseif mode ="temp" then  '�ӽ�����

		if storeid = "" then storeid = ""
		if socname = "" then socname = ""
		if chargeid = "" then chargeid = ""
		if chargename = "" then chargename = ""
		if comment = "" then comment = ""
		if scheduledt = "" then scheduledt =  NULL

	puserdiv=""
	sqlStr = " select top 1 isNULL(userdiv,'') as puserdiv" & vbcrlf	' ���ó üũ�߰�	' 2019.05.23 �ѿ�� �߰�
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner" & vbcrlf
	sqlStr = sqlStr & " where id='" + storeid + "'" & vbcrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		puserdiv = rsget("puserdiv")
	else
		response.write "<script type='text/javascript'>alert('�ش� �Ǵ� ���ó�� �����ϴ�.');</script>"
		response.write "<script type='text/javascript'>location.replace('"& refer &"')</script>"
		rsget.close : dbget.close() : response.End
	end if
	rsget.close

	if masterid = "" or masterid = "0" then
		'��ü�� �˻�
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + storeid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			socname = rsget("socname_kor")
		end if
		rsget.close

		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		rsget("socid") = storeid
		rsget("socname") = socname
		rsget("chargeid") = chargeid
		rsget("chargename") = chargename
		rsget("divcode") = divcode
		rsget("vatcode") = "008"
		rsget("comment") = comment
		rsget("scheduledt") = scheduledt
		rsget("statecd") = 0
		'rsget("executedt") = executedt

		if (Left(storeid,10)="streetshop") or (Left(storeid,9)="wholesale") or (puserdiv="501") or (puserdiv="502") or (puserdiv="503") then
			rsget("ipchulflag") = "S"
		else
			rsget("ipchulflag") = "E"
		end if


		rsget.update
		iid = rsget("id")
		rsget.close

		code = "SO" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + code + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)
		rsget.Open sqlStr,dbget,1

		sqlStr = " delete from [db_storage].[dbo].tbl_acount_storage_detail where mastercode= '" + CStr(code) + "'"
		dbget.execute sqlStr
	ELSE
			iid = masterid
	end if

	'''2.�¶��� �԰� ������ �Է�
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemname = trim(itemnamearr(i))
			itemoptionname = trim(itemoptionnamearr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))

			'���� ������ �پ��� ���̴�.
			itemno = -1 * CInt(trim(itemnoarr(i)))

			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + code + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1
		end if
	next

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,socid='"+storeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,socname='"+socname+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargeid='"+chargeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargename='"+chargename+"'" + VBCrlf
	sqlStr = sqlStr + " ,divcode='"+divcode+"'" + VBCrlf
	sqlStr = sqlStr + " ,vatcode='008'" + VBCrlf
	sqlStr = sqlStr + " ,comment='"+comment+"'" + VBCrlf
	IF not isNull(scheduledt) then
	sqlStr = sqlStr + " ,scheduledt='"+scheduledt+"'" + VBCrlf
	end if
	sqlStr = sqlStr + " ,statecd=0" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(code) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('��� ������ �ӽ�����Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=" + storeid + "&statecd=1&menupos="+menupos+"')</script>"
	dbget.close()	:	response.End

elseif mode="delete" then
	''����� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.code, m.executedt, m.ipchulflag, m.deldt, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)
	rsget.Open sqlStr, dbget, 1

	if statecd<>"0" and not(isnull(statecd)) then
		' �����α�����		' 2021.03.09 �ѿ��
		chulgo_edit_log masterid,finishid,"��ü����"
	end if

	if (not isdeleted) then
		''QuickUpdateNewIpgoDetailSummary code, true

		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
		dbget.Execute sqlStr, AssignedRows

		if (AssignedRows>0) then
		    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
		end if

	end if

	response.write "<script>alert('���� �Ǿ����ϴ�.')</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=&statecd=1')</script>"
	dbget.close()	:	response.End

elseif mode="chulgo" then

	''����� - ���Ȱ��� ����� ���� �ʴ´�.
	sqlStr = "select top 1  m.code, m.executedt, socid as shopid "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		shopid = rsget("shopid")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set executedt='" + executedt + "' , statecd = 7, finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)

	dbget.Execute(sqlStr)

	' �����α�����		' 2021.03.09 �ѿ��
	chulgo_edit_log masterid,finishid,"���Ȯ��"

	if (yyyymmdd<>"") then
		response.write "<script>alert('Err - �̹� ���ó�� �� �����Դϴ�.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	end if

	''���� ���� �����Ǹż���

	'' item
	sqlstr = " update [db_item].[dbo].tbl_item"
	sqlstr = sqlstr + " set limitsold=limitsold - T.chulno"
	sqlstr = sqlstr + " from "
	sqlstr = sqlstr + " ("
	sqlstr = sqlstr + " 	select d.itemid, sum(d.itemno) as chulno"
	sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlstr = sqlstr + " 	where d.mastercode = '" + code + "'"
	sqlstr = sqlstr + " 	and d.deldt is NULL"
	sqlstr = sqlstr + " 	and d.itemno<0"
	sqlstr = sqlstr + " 	and d.iitemgubun='10'"
	sqlstr = sqlstr + " 	group by d.itemid"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"
	sqlstr = sqlstr + " and [db_item].[dbo].tbl_item.limityn='Y'"

	dbget.Execute(sqlStr)

	''�ɼ��ִ»�ǰ
	sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
	sqlStr = sqlStr + " set optlimitsold=optlimitsold - T.chulno" + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlstr = sqlstr + " ("
	sqlstr = sqlstr + " 	select d.itemid, d.itemoption, sum(d.itemno) as chulno"
	sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlstr = sqlstr + " 	where d.mastercode = '" + code + "'"
	sqlstr = sqlstr + " 	and d.deldt is NULL"
	sqlstr = sqlstr + " 	and d.itemno<0"
	sqlstr = sqlstr + " 	and d.iitemgubun='10'"
	sqlstr = sqlstr + " 	group by d.itemid, d.itemoption"
	sqlstr = sqlstr + " ) as T"
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

	dbget.Execute(sqlStr)

	'// ��ո��԰� => ��������԰� '' �ּ�ó��..2017/04/07 ��Ÿ����� �ݿ��ؾ���?
	'// ��Ÿ���+������� �� �ݿ�, ���ν������� ó��, skyer9
	sqlStr = " exec [db_storage].[dbo].[usp_Ten_AvgIpgoPriceToAccoundStorageBuycash] '" & code & "' "
	dbget.Execute sqlStr

	''���ݿ�
	''QuickUpdateNewIpgoDetailSummary code,false
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	'// ������� �ݿ�
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & code & "', 'N' "
	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

	response.write "<script>alert('���ó�� �Ǿ����ϴ�..')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"
	dbget.close()	:	response.End

elseif mode="chulgo2jupsu" Then
	'// ���Ϸ����� üũ(�����̵� ���ϱ�)
	sqlStr = "select top 1  m.code, m.executedt, socid as shopid "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		shopid = rsget("shopid")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""

	if (yyyymmdd = "") then
		response.write "<script>alert('Err - ������� �����Դϴ�.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	else
		'// �������·� ��ȯ, ������� ����
		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
		sqlStr = sqlStr + " set executedt=NULL , statecd = 1, finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
		sqlStr = sqlStr + " where id=" + CStr(masterid)
		dbget.Execute(sqlStr)
	end if

	'// ���ݿ�(����) : ������ ����
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	'// ���ݿ�(����) : ������ ����
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & code & "', 'N' "
	dbget.Execute sqlStr

	' �����α�����		' 2021.03.09 �ѿ��
	chulgo_edit_log masterid,finishid,"������ȯ"

	response.write "<script>alert('������� ��ȯ �Ǿ����ϴ�..')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"
	dbget.close()	:	response.End

elseif mode="chchulgodate" Then
	''����� - ������� ������� ����
	sqlStr = "select top 1  m.code, m.executedt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	if (yyyymmdd = "") then
		response.write "<script>alert('Err - ������� �����Դϴ�.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	end If

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set executedt='" + executedt + "', finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)
	dbget.Execute(sqlStr)

	sqlStr = " update "
	sqlStr = sqlStr + " [db_storage].[dbo].[tbl_ordersheet_master] "
	sqlStr = sqlStr + " set beasongdate = '" + executedt + "', ipgodate = '" + executedt + "', finishuser = '" & finishid & "', finishname = '" & finishname & "' "
	sqlStr = sqlStr + " where alinkcode = '" + code + "' and statecd = '7' "
	dbget.Execute(sqlStr)

	''���ݿ�
	''QuickUpdateNewIpgoDetailSummary code,false
    ''sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	sqlStr = " exec db_summary.dbo.[sp_Ten_recentIpChul_changeChulgoDate] '" & code & "','" & yyyymmdd & "','" & executedt & "' "
	dbget.Execute sqlStr

	' �����α�����		' 2021.03.09 �ѿ��
	chulgo_edit_log masterid,finishid,"������ں���"

    response.write "<script>alert('����� �ݿ��Ǿ����ϴ�.')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"

elseif mode="editmaster" then

	''����� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.code, m.executedt, m.ipchulflag, m.deldt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	if (yyyymmdd<>CStr(executedt)) and (not isdeleted) then
		QuickUpdateNewIpgoDetailSummary code, true
	end if

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set updt = getdate() " + VbCrlf
	if (executedt <> "") then
		sqlStr = sqlStr + " ,executedt='" + executedt + "' " + VbCrlf
	end if
	sqlStr = sqlStr + " ,comment='" + comment + "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)

	dbget.Execute(sqlStr)

	' �����α�����		' 2021.03.09 �ѿ��
	chulgo_edit_log masterid,finishid,"�����,�ڸ�Ʈ����"

	''���ݿ�
	if (yyyymmdd<>CStr(executedt)) and (code<>"") and (not isdeleted)  then
		''QuickUpdateNewIpgoDetailSummary code, false
		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'','" & yyyymmdd & "'"
    	dbget.Execute sqlStr, AssignedRows

    	if (AssignedRows>0) then
    	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
    	end if

	end if

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="deldetail" then
	iid = request("masterid")
	chk= request("chk") + ",,"
	chk = split(chk, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	''����� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	for i=0 to UBound(chk) - 1
		tmpitemgubun = ""
		tmpitemid = ""
		tmpitemoption = ""
		tmpitemno = 0

		if (trim(chk(i)) <> "") then

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set deldt=getdate()" + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))
			dbget.Execute(sqlStr)

			''��� �ݿ�
''			if (tmpitemgubun<>"") then
''				if ipchulflag="S" then
''					QuickUpdateItemChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, tmpitemno*-1,(tmpitemno>0)
''				elseif ipchulflag="E" then
''					QuickUpdateItemEtcChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, tmpitemno*-1,(tmpitemno>0)
''				else
''					'' Nothing
''				end if
''			end if
		end if
	next

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute(sqlStr)

	if statecd<>"0" and not(isnull(statecd)) then
		' �����α�����		' 2021.03.09 �ѿ��
		chulgo_edit_log iid,finishid,"��ǰ����"
	end if

    '' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="editdetail" then
	iid = request("masterid")
	alinkcode = request("alinkcode")
	currencyUnit = request("currencyUnit")
	currencyUnit_Pos = request("currencyUnit_Pos")
	priceChange = 0

	chk= request("chk") + ",,"
	chk = split(chk, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	sellcash = request("sellcash") + ",,"
	sellcash = split(sellcash, ",")

	buycash = request("buycash") + ",,"
	buycash = split(buycash, ",")

	suplycash = request("suplycash") + ",,"
	suplycash = split(suplycash, ",")

	''����� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if

	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	'�ֹ����� ���� ��� �ֹ��� ���� ������ ���� ����
	if alinkcode<>"" then
		if currencyUnit ="KRW" then
			priceChange=1
		end if
		if currencyUnit_Pos ="KRW" then
			priceChange=2
		end if
	end if
	if priceChange > 0 then
		''�ֹ��ڵ�� masteridx�� ã��
		sqlStr = "select top 1 idx, isnull(sitename,'') as sitename"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master"
		sqlStr = sqlStr + " where baljucode='" & Cstr(alinkcode) & "'"

		rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			masteridx = rsget("idx")
			sitename = rsget("sitename")
		end if
		rsget.close
		if sitename<>"WSLWEB" then priceChange=0
	end if

	for i=0 to UBound(chk) - 1
		'response.write  trim(chk(i))
		if (trim(chk(i)) <> "") then
			tmpitemgubun = ""
			tmpitemid = ""
			tmpitemoption = ""
			tmpitemno = 0

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close


			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(itemno(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))
			dbget.Execute(sqlStr)

			if priceChange>0 then

				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + VBCrlf
				sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
				sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,realitemno=" + CStr(itemno(CInt(chk(i)))*-1) + " " + VBCrlf
				if priceChange>1 then
				sqlStr = sqlStr + " ,foreign_sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,foreign_suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
				end if
				sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + " " + VBCrlf
				sqlStr = sqlStr + " and itemid=" + CStr(tmpitemid) + " " + VBCrlf
				sqlStr = sqlStr + " and itemoption='" + CStr(tmpitemoption) + "'"

				'response.write sqlStr & "<Br>"
				'response.end
				dbget.Execute(sqlStr)
			end if

			''��� �ݿ�
''			if (yyyymmdd<>"") and (tmpitemgubun<>"") and (CStr(tmpitemno)<>CStr(itemno(CInt(chk(i))))) then
''				if ipchulflag="S" then
''					QuickUpdateItemChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, (itemno(CInt(chk(i)))-tmpitemno),(tmpitemno>0)
''				elseif ipchulflag="E" then
''					QuickUpdateItemEtcChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, (itemno(CInt(chk(i)))-tmpitemno),(tmpitemno>0)
''				else
''					'' Nothing
''				end if
''			end if
		end if
	next

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute(sqlStr)

	'�ֹ� ������ �հ� �ݾ� ����
	if priceChange>0 then
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf		'/���ֽ� �ؿ� �Һ��ڰ�
		sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'/���ֽ� �ؿ� ���ް�
		sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'/Ȯ�� �ؿ� �Һ��ڰ�
		sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'/Ȯ�� �ؿ� ���ް�
		sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
		sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T" + vbCrlf
		sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
	end if

	if statecd<>"0" and not(isnull(statecd)) then
		' �����α�����		' 2021.03.09 �ѿ��
		chulgo_edit_log iid,finishid,"��ǰ����"
	end if

    '' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

	response.write "<script>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="adddetail" then
	iid = request("masterid")

	''����� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	'''2.�¶��� �԰� ������ �߰�
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))
			itemno = CInt(trim(itemnoarr(i))) * -1
			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = ""
			itemoptionname = ""

			''sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			''sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			''sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			''sqlStr = sqlStr + " values('" + request("code") + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			''rsget.Open sqlStr,dbget,1

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + request("code") + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1

			''��� �ݿ�
''			if ipchulflag="S" then
''				QuickUpdateItemChulgoSummary  yyyymmdd, itemgubun, itemid, itemoption, itemno,(itemno>0)
''			elseif ipchulflag="E" then
''				QuickUpdateItemEtcChulgoSummary  yyyymmdd, itemgubun, itemid, itemoption, itemno,(itemno>0)
''			else
''				'' Nothing
''			end if

		end if
	next

	'// ��ǰ���� ���� ������ ������Ʈ, skyer9, 2016-11-28
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	if statecd<>"0" and not(isnull(statecd)) then
		' �����α�����		' 2021.03.09 �ѿ��
		chulgo_edit_log masterid,finishid,"��ǰ�߰�"
	end if

    '' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

	response.write "<scr" + "ipt>location.replace('" + refer + "')</sc" + "ript>"
	dbget.close()	:	response.End
elseif mode="wichulgoconv" then
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set mwgubun='H'" + vbCrlf
	sqlStr = sqlStr + " where id=" + CStr(request("id"))
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('OK')</script>"
	response.write "<script>window.close()</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
