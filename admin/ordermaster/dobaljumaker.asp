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
' Description : �������
' History : �̻� ����
'			2018.03.28 �ѿ�� ����
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
'��ǰ�����
if workgroup="N" then baljutype="S"
if workgroup="M" then baljutype="S"
if workgroup="J" then baljutype="S"

''��� �ù�� �߰� int
songjangdiv = request("songjangdiv")

dummiserial = orderserial
orderserial = split(orderserial,"|")
sitename    = split(sitename,"|")

if mode="arr" then
	''��ȿ��üũ.
	dummiserial = Mid(dummiserial,2,Len(dummiserial))
	dummiserial = replace(dummiserial,"|","','")
	sqlStr = " select top 1 orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " where orderserial in ('" + dummiserial + "')"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		response.write "<script language='javascript'>"
		response.write "alert('" + rsget("orderserial") + " : �̹� ������õ� �ֹ��Դϴ�. \n\n�̹� ��������� �ֹ����� �ߺ� ������� �� �� �����ϴ�.');"
		response.write "location.replace('" + CStr(refer) + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	'////////////////// ������� ���ⱸ�� �ֹ��и� ////////////////		'2019.02.13 �ѿ��
	'/ �ֹ��� �߿� ���ⱸ�� ��ǰ�� �ִ��� üũ
	sqlStr = "select d.orderserial, so.orgitemid, so.orgitemoption, so.reserveitemid, so.reserveitemoption" & vbcrlf
	sqlStr = sqlStr & " from db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sqlStr = sqlStr & " join db_item.dbo.tbl_item i with (nolock)" & vbcrlf
	sqlStr = sqlStr & " 	on d.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.itemdiv = '75'" & vbcrlf		' ���ⱸ�� ��ǰ
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
			' ���� ��ǰ ��Ī �ȵ� �ֹ���
			if isnull(rsget("reserveitemoption")) or rsget("reserveitemoption")="" then
				notmatchingstandingorder = notmatchingstandingorder & rsget("orderserial") & ","
			Else
				' ���ⱸ���ֹ�
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
		response.write "	alert('[���ⱸ��] �����߼� ��ǰ ��Ī�� �����Ǿ� ���� �ʽ��ϴ�.\n�ش� �ֹ����� ���ܵǰ� ������� �˴ϴ�.\n�ش� �μ��� ���ⱸ�� ��ǰ��Ī ��û�ϼ���.\n\n' + "& notmatchingstandingorder &");"
		response.write "</script>"
	end if

	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			' ���ⱸ�� ��ǰ�� ��� ���� ����
			if instr(standingorderserial,orderserial(i)) > 0 then
				standingorderitemid = ""
				standingorderitemoption = ""
				startreserveidx = ""

				'/ ���ⱸ�� ��ǰ��ȣ�� �ɼǹ�ȣ�� �޾ƿ�
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

				' ��ġ����Ŀ ���ⱸ�� ������� �ֹ��и�
				sqlStr = "exec db_order.[dbo].sp_Ten_StandingOrder_make "& standingorderitemid &",'"& standingorderitemoption &"',"& startreserveidx &",'"& orderserial(i) &"','SCM'"

				'response.write sqlStr & "<br>"
				dbget.execute sqlStr
			end if
		end if
	next
	'////////////////// ������� ���ⱸ�� �ֹ��и� ////////////////

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

    '' �ù�纰 ������� system���� ����

	'#######������� ������###############
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

	'#######������� ������###############
	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			' ���ⱸ�� ���� �߼� ��ǰ�� ��Ī�� �ȵǾ� ������� ����.		'2019.02.13 �ѿ��
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

    ''** [db_order].[dbo].tbl_baljudetail.baljusongjangno is NULL �ΰ�� ��ü������� �ν� (Logics ���� �ý���)
    ''�ٹ����� ����� ��� �����ȣ�� not null ������ �Է�..

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

	'// �ؿܹ��(EMS) �� ������ �ٹ�
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

	'// �ؿܹ��(UPS) �� ������ �ٹ�
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

	''' �ö�� �ֹ��� CASE - ���� Ȯ���� ���Ȱ��;
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.ipkumdiv>4"
	''sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.ipkumdiv<8"  ''�� ��� �Ϸ� �Ȱ͵� �������..
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.baljudate is NULL"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1


	'#######�ֹ� �뺸 ���� ############### (��������� ����)
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set ipkumdiv='5'"
	sqlStr = sqlStr + " ,baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.ipkumdiv=4"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"

	rsget.Open sqlStr,dbget,1


	''' ��ü��� Ȯ���� �� �� ���� ���̿� ���̰� �Ǹ� ����������� �Է��� �ȵǴ� ��찡 �߻��Ѵ�.
	''' ��������� ���Էµ� ��� �ֹ��� ��������� �Է�
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

    '#######  �ù���Է� (Master) ############### (������ ���� �Է����� ���� , �ù�縸 �Է���.)
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
	'###### Order Detail �ٹ����� ��� ��������� ���� ############
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

	'###### Order Detail ��ü ��� ��ü�ֹ��뺸 flag ���� : NULL �� ��츸 ############
    sqlStr = "update [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_detail.orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_order].[dbo].tbl_baljudetail"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.isupchebeasong='Y'"
	sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_detail.itemid<>0"
	sqlStr = sqlStr + " and IsNULL([db_order].[dbo].tbl_order_detail.currstate,0)=0"

	'' ��ü����� ��� ��ü�ֹ��뺸 or NULL �ΰ�� ��� Ȯ�� ������.
	rsget.Open sqlStr,dbget,1
end if


If Err.Number = 0 Then
        errcode = "006"

	''������� ������ ���
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

	''��ǰ���ڵ����� - �������Ȱ͸�

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

	''��ǰ���̺����
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

	''�귣�����̺����
	sqlStr = " update [db_temp].[dbo].tbl_baljuitem" + VbCrlf
	sqlStr = sqlStr + " set rackcode = IsNULL(T.prtidx,0)" + VbCrlf
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c T" + VbCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_baljuitem.baljuid=" + CStr(iid) + VbCrlf
	sqlStr = sqlStr + " and [db_temp].[dbo].tbl_baljuitem.makerid=T.userid" + VbCrlf
	rsget.Open sqlStr,dbget,1
end if

If Err.Number = 0 Then
        errcode = "011"

	''�ɼ����̺����
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
    '' ��� ������Ʈ
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_balju " & iid
    dbget.execute sqlStr
end if

''If Err.Number = 0 Then
''        errcode = "013"
''    ''��ü �̹��/��������� ���� ������Ʈ
''    sqlStr = " exec db_partner.[dbo].sp_Ten_Partner_Summary_Upche_Mibalju_MiBeasong ''"
''
''    dbget.Execute sqlStr
''end if

If Err.Number = 0 Then
    dbget.CommitTrans

    if (Left(pickingStationCd, 3) = "IFC") then
        '// Ʈ������ ���� �Ŀ� �����ؾ� �Ѵ�.
        '' pickingStationCd : IFC2/IFC3/IFC4/IFC5
        '' errcode = "014"
        '' sqlStr = " exec [db_aLogistics].[dbo].[usp_LogisticsItem_Balju_to_interface] " & iid & ", '" & pickingStationCd & "', '" & pickingStationCdArrStr & "', '" & session("ssBctId") & "' "
        '' dbget_Logistics.execute sqlStr
    end if
Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n������ ���� ��� (�����ڵ� : " + CStr(errcode) + ")');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
End If
on error Goto 0



''�߰� �۾�. Ʈ����ǿ��� ��.. �ֹ����� ��� �ð��� �����ɸ� : ����ǰ�ۼ��ð�.
''response.write "<font color=red>������� ������ �������� ��� �����岲 ���ǿ��!! - ����ǰ���� </font><br>"
    '' ������ ���� ����
    sqlStr = " delete from [db_temp].[dbo].tbl_baljuitem " + VbCrlf
    sqlStr = sqlStr + " where baljuid<" + CStr(iid-100)

    dbget.execute sqlStr

''������ �ӽ�
''sqlStr = " exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_IMSI] "& iid
''dbget.execute sqlStr

' õ�鸸�� ����Ʈī�� �̺�Ʈ	' 2018.03.28 �ѿ�� ����
'if (date()>="2018-04-02") and (date()<"2018-04-18") then
'	sqlStr = "exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_GiftCard_TenDLV] " & iid
'
'	'response.write sqlStr & "<br>"
'	dbget.execute sqlStr
'end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''���� ù���� ���ڰ� �Ѵٸ� �ּ� Ǯ�� // 8/21 ������ð����� ���� �Ǿ���.
'if (now()>="2017-08-14") and (now()<"2017-09-02") then
'	sqlStr = " exec [db_order].[dbo].[sp_Ten_order_Gift_BALJU_FirstOrder_TenDLV] " & iid
'	dbget.execute sqlStr
'end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' ����ǰ ��� -> �ֹ��� ����.
''����ǰ �ٹ��
''If Err.Number = 0 Then
''        errcode = "013"
''
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker " & iid & ",'N',0"
''    dbget.execute sqlStr
''end if

''    ''����ǰ ��ü���
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker " & iid & ",'Y',0"
''    dbget.execute sqlStr


''    ���� �̺�Ʈ
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',9098,'10x10'"
''    dbget.execute sqlStr
''response.write "����ǰ �ۼ� (����) �Ϸ�<br>"

''    ''����ǰ ���̾ ����ǰ  �̺�Ʈ;;
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',8752,'10x10'"
''    dbget.execute sqlStr

    ''����ǰ ���̾ ����ǰ  �̺�Ʈ 2��° ���� Case;;
''    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(iid) & ",'N',8842,'10x10'"
''    dbget.execute sqlStr
''response.write "����ǰ �ۼ� (���̾ �̺�Ʈ) �Ϸ�<br>"


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
alert('������ü��� ���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
