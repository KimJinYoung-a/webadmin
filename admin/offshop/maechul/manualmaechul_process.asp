<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ����
' History : 2012.08.07 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim i , mode , bagidxarr , sqlStr ,menupos, shopid ,barcode ,sqlsearch ,shopregdate ,orderno ,posid ,result
dim adminuserid , masteridx ,cnt, nowdate, jungsandate
dim itemgubun, itemid, itemoption, itemprice, suplycash, buyprice, itemname, itemoptionname, makerid, extbarcode
dim itemgubunarr ,itemidarr ,itemoptionarr ,itemnamearr ,itemoptionnamearr ,sellcasharr ,suplycasharr
dim shopbuypricearr ,itemnoarr ,makeridarr ,extbarcodearr
dim imaechulgubun, tmpshopid
    mode = requestcheckvar(request("mode"),32)
    menupos = requestcheckvar(request("menupos"),10)
    shopregdate = requestcheckvar(request("shopregdate"),10)
	shopid = requestcheckvar(request("shopid"),32)
	barcode = requestcheckvar(request("barcode"),32)
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	itemnamearr = request("itemnamearr")
	itemoptionnamearr = request("itemoptionnamearr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	shopbuypricearr = request("shopbuypricearr")
	itemnoarr = request("itemnoarr")
	makeridarr = request("makeridarr")
	extbarcodearr = request("extbarcodearr")

adminuserid = session("ssBctId")
posid = 99
nowdate = now()
jungsandate = year(nowdate) & "-" & Format00(2,month(nowdate)) & "-" & "10"
'response.write mode

'//���ڵ� ��ǰ���
if mode = "oneaddmanualItem" then

	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ �����ϴ�.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if barcode = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���ڵ带 �Է� �ϼ���.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if len(barcode) < 11 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���ڵ��� ���̰� ª���ϴ�.\n�����ڵ峪 ������ڵ带 �ٽ� Ȯ����, �Է� �ϼ���.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	if trim(barcode)<>"" then

		'//���ڵ尡 �������, ������ڵ�� �ʼ��� �˻�
		sqlStr = "select top 1"
		sqlStr = sqlStr + " itemgubun,shopitemid,itemoption"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr + " where extbarcode='" + trim(barcode) + "'"

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			itemgubun = rsget("itemgubun")
			itemid = rsget("shopitemid")
			itemoption = rsget("itemoption")
		end if
		rsget.Close
	end if

	if itemid = "" then
		itemgubun 	= BF_GetItemGubun(barcode)
		itemid 		= BF_GetItemId(barcode)
		itemoption 	= BF_GetItemOption(barcode)
	end if

	sqlsearch = sqlsearch + " and s.itemgubun='"& itemgubun &"'"
	sqlsearch = sqlsearch + " and s.shopitemid="& itemid &""
	sqlsearch = sqlsearch + " and s.itemoption='"& itemoption &"'"

    sqlStr = " select top 1 s.itemgubun, s.shopitemid, s.itemoption, s.extbarcode, s.isusing as itemstatus"
    sqlStr = sqlStr + " , convert(varchar(32),s.regdate,20) as regdate"
	sqlStr = sqlStr + " ,(CASE"
	sqlStr = sqlStr + " 	when s.shopsuplycash = 0 and sd.comm_cd in ('B011','B012')"		'/���԰��� 0 ,������Ź, ��ü��Ź
	sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultmargin,100))/100)"
	sqlStr = sqlStr + " 	else s.shopsuplycash"
	sqlStr = sqlStr + "	end) as shopsuplycash"
	'sqlStr = sqlStr + " , s.shopsuplycash"
	sqlStr = sqlStr + " ,(CASE" & VbCRLF
	sqlStr = sqlStr + " 	when s.shopbuyprice = 0 and sd.comm_cd in ('B011','B012')"		'/������� 0 ,������Ź, ��ü��Ź
	sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultsuplymargin,100))/100)"
	sqlStr = sqlStr + "		else s.shopbuyprice"
	sqlStr = sqlStr + "	end) as shopbuyprice"
	'sqlStr = sqlStr + " , s.shopbuyprice"
    sqlStr = sqlStr + " , (CASE WHEN s.itemgubun='80' THEN 0 ELSE s.orgsellprice END) as orgsellprice"
    sqlStr = sqlStr + " , (CASE WHEN s.itemgubun='80' THEN 0 ELSE s.shopitemprice END) as shopitemprice"      ''�ǸŰ�
    sqlStr = sqlStr + " , s.makerid ,s.extbarcode" '' �귣�� ID
    sqlStr = sqlStr + " , s.shopitemname, s.shopitemoptionname"
    sqlStr = sqlStr + " , c.socname_kor"        '' �귣�� ��
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
	sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
	sqlStr = sqlStr + " 	on sd.shopid='"&shopid&"' and s.makerid=sd.makerid" & VbCRLF
	sqlStr = sqlStr + " join [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + " 	on s.makerid=c.userid"
	sqlStr = sqlStr + " where 1=1 " & sqlsearch

	'response.write sqlStr & "<Br>"
	rsget.open sqlStr,dbget,1

    if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("shopitemid")
		itemoption = rsget("itemoption")
		itemprice = rsget("shopitemprice")
		suplycash = rsget("shopsuplycash")
		buyprice = rsget("shopbuyprice")
		itemname = rsget("shopitemname")
		itemoptionname = rsget("shopitemoptionname")
		makerid = rsget("makerid")
		extbarcode = rsget("extbarcode")
    end if

    rsget.close

	if itemid <> "" then
		response.write "<script type='text/javascript'>"
		response.write "	opener.ReActItems('"&itemgubun&"|','"&itemid&"|','"&itemoption&"|','"&itemprice&"|','"&suplycash&"|','"&buyprice&"|','1','"&itemname&"|','"&itemoptionname&"|','"&makerid&"|','"&extbarcode&"|');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('�ش�Ǵ� ��ǰ�� �����ϴ�.');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

'//��������
elseif mode = "addmanualItem" then

	if not(C_ADMIN_USER) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ �����ϴ�');"
		response.write "	self.close();"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ �����ϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if shopregdate = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���⳯¥�� �������� �ʾҽ��ϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	itemgubunarr = split(itemgubunarr,"|")
	itemidarr	= split(itemidarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	itemnamearr		= split(itemnamearr,"|")
	itemoptionnamearr = split(itemoptionnamearr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	makeridarr = split(makeridarr,"|")
	extbarcodearr = split(extbarcodearr,"|")

	'//�δ� ���� ���� �ϰ��
	if datediff("m", Left(shopregdate,10) , nowdate) >= 2 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�δ� ���������� �Է� �ϽǼ� �����ϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	'//������ ���� �Է½�
	if datediff("m", Left(shopregdate,10) , nowdate) < 0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ ������ �Է��� �Ұ��� �մϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

'	'//������ ���� �Է½�
'	if datediff("m", Left(shopregdate,10) , nowdate) = 1 then
'		if datediff("d", jungsandate , nowdate) > 0 then
'			response.write "<script type='text/javascript'>"
'			response.write "	alert('�������� ������ ������ �� ��¥ �Դϴ�.');"
'			response.write "</script>"
'			if Not C_ADMIN_AUTH then
'				response.end	:	dbget.close()
'			else
'				response.write "<script type='text/javascript'>"
'				response.write "	alert('[�����ڱ���]\n\n��������մϴ�.');"
'				response.write "</script>"
'			end if
'		end if
'	end if

	cnt = UBound(itemidarr)

	for i=0 to cnt - 1

		'//�ǸŰ��� ���� �Ѵ� ���̳ʽ� �ϰ��..���ϸ� �÷����� ����.. �ðܳ�
		if left(trim(sellcasharr(i)),1)="-" and left(trim(itemnoarr(i)),1)="-" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�ǸŰ��� ���� �Ѵ� ���̳ʽ� ���� �ɼ� �����ϴ�.\n���̳ʽ� �ֹ� �Է½� ������ ���̳ʽ��� �Է����ּ���');"
			response.write "</script>"
			response.end	:	dbget.close()
		end if
	next

	orderno = manualordernomake_off(shopid,posid)

    '/�̹������ϴ� �ֹ���ȣ���� üũ
    sqlStr = "select count(idx) as cnt"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
	sqlStr = sqlStr + " where orderno='"&orderno&"'"

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1

	if Not rsget.Eof then
	    if (rsget("cnt")>0) then result = "Y"
	end if

	rsget.close

	if result = "Y" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�ֹ���ȣ�� �̹� ���� �մϴ�. ������ ���ǿ��.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	result = ""

    ''���ⱸ��  /2013/12/17 �߰�
    imaechulgubun=""

    sqlStr = "select isNULL(tplcompanyid,'MANUAL') as maechulgubun"
    sqlStr = sqlStr&" from db_partner.dbo.tbl_partner"
    sqlStr = sqlStr&" where id='"&shopid&"'"
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		imaechulgubun=rsget("maechulgubun")
	end if
	rsget.close

    if (imaechulgubun="") then
        imaechulgubun="MANUAL"
    end if

	'// �Է��� ����Ÿ ����
	'// 1. �ùٸ� ���ڵ�����
	'// 2. ����� �ִ���
	for i=0 to cnt - 1

		sqlStr = " 	select top 1 i.itemgubun, IsNull(s.shopid, '') as shopid " + vbcrlf
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid='"&shopid&"'"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii" & VbCRLF
		sqlStr = sqlStr + " 		on i.shopitemid = ii.itemid" & VbCRLF
		sqlStr = sqlStr + " 		and i.itemgubun = '10'" & VbCRLF
		sqlStr = sqlStr + " 	where i.itemgubun = '"& requestCheckVar(trim(itemgubunarr(i)),2) &"'" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid = "& requestCheckVar(trim(itemidarr(i)),10) &"" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption = '"& requestCheckVar(trim(itemoptionarr(i)),4) &"'" + vbcrlf
		tmpshopid = "XXXXXXXXX"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			tmpshopid = rsget("shopid")
		end if
		rsget.close

		if (tmpshopid = "XXXXXXXXX") then
			response.write "�߸��� ��ǰ�ڵ� �Ǵ� �������� ��ǰ��� ���� ��ǰ�Դϴ�. : " & itemgubunarr(i)
			dbget.close() : response.end
		elseif tmpshopid = "" then
			response.write "����� �������� �ʾҽ��ϴ�. : " & itemgubunarr(i)
			dbget.close() : response.end
		end if
	next

	'//������ ���̺� ���
    sqlStr = "select * from [db_shop].[dbo].tbl_shopjumun_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("orderno")    = orderno
	rsget("shopid")     = shopid
	rsget("totalsum")   = 0
	rsget("realsum")    = 0
	rsget("jumundiv")   = "00"
	rsget("jumunmethod") = "01"
	rsget("shopregdate") = Left(shopregdate,10)
	rsget("cancelyn")   = "N"
	rsget("shopidx")    = "0"
	rsget("spendmile")  = "0"
	rsget("pointuserno") = ""
	rsget("gainmile") = "0"
	rsget("cashsum")    = 0
    rsget("cardsum")    = "0"
    rsget("casherid")   = adminuserid
    rsget("GiftCardPaySum") = "0"
    rsget("CardAppNo")      = ""
    rsget("CashReceiptNo")  = ""
    rsget("CashreceiptGubun") = ""
    rsget("CardInstallment")  = ""
	rsget("IXyyyymmdd") = Left(shopregdate,10)
	rsget("tableno")  = "0"
    rsget("TenGiftCardPaySum")  = "0"
	rsget("TenGiftCardMatchCode")  = ""
	rsget("refOrderNo")  = ""
	rsget("maechulgubun")  = imaechulgubun '"MANUAL"

	rsget.update
		masteridx = rsget("idx")
	rsget.close

	for i=0 to cnt - 1

		'//������ ���̺� ���
        sqlStr = "insert into [db_shop].[dbo].tbl_shopjumun_detail" + vbcrlf
		sqlStr = sqlStr + " ( masteridx, orderno, itemgubun, itemid, itemoption" + vbcrlf
		sqlStr = sqlStr + " , itemno, itemname, itemoptionname, sellprice, realsellprice" + vbcrlf
		sqlStr = sqlStr + " , suplyprice" + vbcrlf
		sqlStr = sqlStr + " , shopbuyprice" + vbcrlf
		sqlStr = sqlStr + " , makerid, jungsanid, cancelyn" + vbcrlf
		sqlStr = sqlStr + " , shopidx, itempoint, discountKind, Iorgsellprice, Ishopitemprice" + vbcrlf
		sqlStr = sqlStr + " , jcomm_cd, addtaxcharge, vatinclude)" + vbcrlf
		sqlStr = sqlStr + " 	select" + vbcrlf
		sqlStr = sqlStr + " 	'"&masteridx&"','"&orderno&"',i.itemgubun ,i.shopitemid ,i.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	,'"& requestCheckVar(trim(itemnoarr(i)),10) &"', i.shopitemname, i.shopitemoptionname,'"& requestCheckVar(trim(sellcasharr(i)),20) &"','"& requestCheckVar(trim(sellcasharr(i)),20) &"'" + vbcrlf
		sqlStr = sqlStr + " 	,(CASE" & VbCRLF
		sqlStr = sqlStr + " 		when isnull(ii.mwdiv,'')='M' and s.comm_cd not in ('B012')" & VbCRLF		'//�¶��θ����̰�, ��ü��Ź�� �ƴϸ� �¶��θ��԰���
		sqlStr = sqlStr + " 			THEN isnull(ii.buycash,0)" & VbCRLF
		'sqlStr = sqlStr + " 		when i.shopsuplycash = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/���԰��� 0 ,������Ź, ��ü��Ź ,�����Ź
		sqlStr = sqlStr + " 		when i.shopsuplycash = 0" & VbCRLF		'���԰� �� ������ ������
		sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultmargin,100))/100)" & VbCRLF
		sqlStr = sqlStr + " 		else i.shopsuplycash" & VbCRLF
		sqlStr = sqlStr + "			end) as shopsuplycash" & VbCRLF
		sqlStr = sqlStr + " 	,(CASE" & VbCRLF
		'sqlStr = sqlStr + " 		when i.shopbuyprice = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/������� 0 ,������Ź, ��ü��Ź ,�����Ź
		sqlStr = sqlStr + " 		when i.shopbuyprice = 0" & VbCRLF		'������� �� ������ ������
		sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultsuplymargin,100))/100)" & VbCRLF
		sqlStr = sqlStr + "			else i.shopbuyprice" & VbCRLF
		sqlStr = sqlStr + "			end) as shopbuyprice" & VbCRLF
		sqlStr = sqlStr + " 	, i.makerid, i.makerid, 'N'" + vbcrlf
		sqlStr = sqlStr + " 	,'0','0','0', i.orgsellprice, i.shopitemprice" + vbcrlf
		sqlStr = sqlStr + " 	, s.comm_cd, 0, i.vatinclude" + vbcrlf
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid='"&shopid&"'"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii" & VbCRLF
		sqlStr = sqlStr + " 		on i.shopitemid = ii.itemid" & VbCRLF
		sqlStr = sqlStr + " 		and i.itemgubun = '10'" & VbCRLF
		sqlStr = sqlStr + " 	where i.itemgubun = '"& requestCheckVar(trim(itemgubunarr(i)),2) &"'" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid = "& requestCheckVar(trim(itemidarr(i)),10) &"" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption = '"& requestCheckVar(trim(itemoptionarr(i)),4) &"'" + vbcrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr
	next

	'//������ ���̺� �ջ�
	sqlStr = "update m" + vbcrlf
	sqlStr = sqlStr + " set m.totalsum = t.sellprice" + vbcrlf
	sqlStr = sqlStr + " ,m.realsum = t.realsellprice" + vbcrlf
	sqlStr = sqlStr + " ,m.cashsum = t.realsellprice" + vbcrlf
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
	sqlStr = sqlStr + " join (" + vbcrlf
	sqlStr = sqlStr + " 	select" + vbcrlf
	sqlStr = sqlStr + " 	orderno ,sum((d.sellprice+addtaxcharge) * d.itemno) as sellprice" + vbcrlf
	sqlStr = sqlStr + " 	,sum((d.realsellprice+addtaxcharge) * d.itemno) as realsellprice" + vbcrlf
	sqlStr = sqlStr + " 	,sum((d.suplyprice+addtaxcharge) * d.itemno) as suplyprice" + vbcrlf
	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
	sqlStr = sqlStr + " 	where d.cancelyn = 'N'" + vbcrlf
	sqlStr = sqlStr + " 	and d.orderno = '"&orderno&"'" + vbcrlf
	sqlStr = sqlStr + " 	group by orderno" + vbcrlf
	sqlStr = sqlStr + " ) as t" + vbcrlf
	sqlStr = sqlStr + " 	on m.orderno = t.orderno" + vbcrlf
	sqlStr = sqlStr + " 	and m.cancelyn = 'N'" + vbcrlf
	sqlStr = sqlStr + " where m.orderno = '"&orderno&"'"

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'// �ߺ��Է� ����
	sqlStr = "[db_shop].[dbo].[usp_TEN_Shop_ManualOrder_DuppRemove] '" & orderno & "'"
	dbget.Execute sqlStr

	''��� ������Ʈ(No tran)
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_RegOrder '" & orderno & "'"

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	parent.self.close();"
	response.write "</script>"
	response.end	:	dbget.close()

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�����ڰ� �����ϴ�.');"
	response.write "</script>"
	response.end	:	dbget.close()
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
