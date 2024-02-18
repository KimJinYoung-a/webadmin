<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%

dim oordermaster, oorderdetail

dim id, divcd, gubun01, gubun02, orderserial, customername, userid, writeuser, finishuser, title, contents_jupsu, contents_finish
dim currstate, regdate, finishdate, refundrequire, refundresult, songjangno, beasongdate, cause, causedetail, requireupche, makerid, deleteyn
dim extsitename, rebankname, rebankaccount, rebankownername, refundbeasongpay, refunditemcostsum, refunddeliverypay, refundadjustpay, returnmethod

dim detailitemlist, detailitemnolist
dim did, dmasterid, dorderserial, ditemid, ditemoption, dmakerid, ditemname, ditemoptionname
dim dregitemno, dconfirmitemno, ditemcost, dbuycash, disupchebeasong, dregdetailstate, dcausediv, dcausedetail, dcausecontent, dcurrstate

dim canceldetailno, canceldetailnamelist

dim sitename
dim canceldetailall
dim refundcstitle


dim mode, i, j, k, tmp
dim sqlStr

mode = html2db(request("mode"))

id              = html2db(RequestCheckvar(request("id"),10))
divcd           = html2db(RequestCheckvar(request("divcd"),10))
gubun01         = html2db(RequestCheckvar(request("gubun01"),10))
gubun02         = html2db(RequestCheckvar(request("gubun02"),10))
orderserial     = html2db(RequestCheckvar(request("orderserial"),16))
customername    = html2db(RequestCheckvar(request("customername"),32))
userid          = html2db(RequestCheckvar(request("userid"),32))
writeuser       = html2db(RequestCheckvar(request("writeuser"),32))
finishuser      = html2db(RequestCheckvar(request("finishuser"),32))
title           = html2db(RequestCheckvar(request("title"),128))
contents_jupsu  = html2db(request("contents_jupsu"))
contents_finish = html2db(request("contents_finish"))
currstate       = html2db(RequestCheckvar(request("currstate"),10))
regdate         = html2db(RequestCheckvar(request("regdate"),32))
finishdate      = html2db(RequestCheckvar(request("finishdate"),32))
refundrequire   = html2db(RequestCheckvar(request("refundrequire"),10))
refundresult    = html2db(RequestCheckvar(request("refundresult"),10))
songjangno      = html2db(RequestCheckvar(request("songjangno"),16))
beasongdate     = html2db(RequestCheckvar(request("beasongdate"),32))
cause           = html2db(RequestCheckvar(request("causecd"),10))
causedetail     = html2db(request("causedetail"))
requireupche    = html2db(RequestCheckvar(request("requireupche"),2))
makerid         = html2db(RequestCheckvar(request("makerid"),32))
deleteyn        = html2db(RequestCheckvar(request("deleteyn"),2))
extsitename     = html2db(RequestCheckvar(request("extsitename"),32))
rebankname              = html2db(RequestCheckvar(request("rebankname"),32))
rebankaccount           = html2db(RequestCheckvar(request("rebankaccount"),32))
rebankownername         = html2db(RequestCheckvar(request("rebankownername"),32))
refundbeasongpay        = html2db(RequestCheckvar(request("refundbeasongpay"),10))
refunditemcostsum       = html2db(RequestCheckvar(request("refunditemcostsum"),10))
refunddeliverypay       = html2db(RequestCheckvar(request("refunddeliverypay"),10))
refundadjustpay         = html2db(RequestCheckvar(request("refundadjustpay"),10))
returnmethod            = html2db(RequestCheckvar(request("returnmethod"),32))
sitename            	= html2db(RequestCheckvar(request("sitename"),32))


detailitemlist 			= html2db(request("detailitemlist"))
detailitemnolist 		= html2db(request("detailitemnolist"))

response.write "divcd=" & divcd & "<br>"
response.write "gubun01=" & gubun01 & "<br>"
response.write "gubun02=" & gubun02 & "<br>"
'dbget.close()	:	response.End
if contents_jupsu <> "" then
	if checkNotValidHTML(contents_jupsu) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end If
if contents_finish <> "" then
	if checkNotValidHTML(contents_finish) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end If
if causedetail <> "" then
	if checkNotValidHTML(causedetail) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end if


if (Len(divcd)=1) then divcd="A00" & divcd
if (Len(divcd)=2) then divcd="A0" & divcd

'==============================================================================
set oordermaster = new CRequestLecture
oordermaster.FRectOrderSerial = orderserial
oordermaster.GetRequestLectureMasterOne

set oorderdetail = new CRequestLecture
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.CRequestLectureDetailList

'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = oordermaster.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture

'==============================================================================
'���üũ
'��ҵ� ���¿� ���� ������� �Ұ���
'��ҺҰ���(���½����� -3��)�� �������� üũ�Ѵ�. >> �����ϰ� ����
'TODO : ��ü����� ���� ���ȯ�ҳ����� �ִٸ� ���� ������ȯ������������ �������ݾ׸� ȯ���Ҽ� �ֵ��� ����
'��ü��ҵ� ��û�� �κ���Ұ� �Ұ����մϴ�.
'ī��/�ǽð���ü ��Ҵ� ���½�û�� ��ҵǾ�߸� �����մϴ�.
'
if ((mode = "cancelorder") and (oordermaster.FOneItem.Fcancelyn = "Y")) then
    response.write "<script>alert('�̹� ��ҵ� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if ((mode = "cancelitem") and (oordermaster.FOneItem.Fcancelyn = "Y")) then
    response.write "<script>alert('�̹� ��ҵ� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

'if ((mode = "cancelorder") and (Left(DateAdd("d",3,now), 10)  > Left(olecture.FOneItem.Flec_startday1, 10))) then
'    response.write "<script>alert('������Ҵ� ���½��� 3�������� �����մϴ�.'); history.back();</script>"
'    dbget.close()	:	response.End
'end if

if ((mode = "cancelcard") and (oordermaster.FOneItem.Fcancelyn <> "Y")) then
    response.write "<script>alert('���½�û�� ��ҵ��� �ʾҽ��ϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if ((mode = "revalidate") and ((oordermaster.FOneItem.Fipkumdiv <> "2") and (oordermaster.FOneItem.Faccountdiv <> "7") and (oordermaster.FOneItem.Fcancelyn <> "Y"))) then
    response.write "<script>alert('��ҵ� �ֹ��� �ƴϰų�, �������ֹ��� �ƴϰų�, �ֹ��������°� �ƴմϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if (mode = "revalidate") then
    '��ȸ���� �ƴѰ��(���� Ȯ��)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '������� Ȯ��
    	    sqlStr = " select top 1 orderserial "
    	    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
    	    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
    	    rsget.Open sqlStr,dbget,1
    	    i = "N"
    	    if  not rsget.EOF  then
    	        i = "Y"
    	    end if
    	    rsget.close

    	    if (i = "N") then
                response.write "<script>alert('�ֹ��� ���Ǿ��� ������ �������� �ʽ��ϴ�. ���ֹ� �ϼ���.'); history.back();</script>"
                dbget.close()	:	response.End
    	    end if
        end if
    end if
end if

if (sitename = "diyitem") then
	refundcstitle = "ȯ�ҿ�û(DIY��ǰ)"
else
	refundcstitle = "ȯ�ҿ�û(����)"
end if

'==============================================================================
if (mode = "cancelorder") then
    '��ü���
    ' - �Ա��� ����� ���, �ֹ����AS��Ϲ�����ó��, �ֹ������, ���ϸ�����ο�, �������������ǥ��, ���� �������� ������
    ' - �Ա��� ����� ���, �ֹ����AS��Ϲ�����ó��, �ֹ������, ���ϸ�����ο�, �������������ǥ��, ���� �������� ������, ȯ�ұݾ��� �������(�ſ�ī����� �Ǵ� �������Ա�AS����������� �̵�)

    '======================================================================
    '�ֹ����AS��Ϲ�����ó��(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
	'rsget("refundrequire")  = 0
	'rsget("cause")          = cause
	'rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = rebankname
	'rsget("rebankaccount")  = rebankaccount
	'rsget("rebankownername")        = rebankownername
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close

	sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " set finishdate=getdate() "
	sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	sqlStr = sqlStr + " ,contents_finish = '" + html2db("�������") + "' "
	sqlStr = sqlStr + " ,currstate = 'B007' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	sqlStr = sqlStr + " where id=" + CStr(id) + " "
	rsget.Open sqlStr,dbget,1

	'======================================================================
	'�ֹ������
	sqlStr = " update [db_academy].[dbo].tbl_academy_order_master "
	sqlStr = sqlStr + " set cancelyn='Y' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
	rsAcademyget.Open sqlStr,dbAcademyget,1

	'======================================================================
	''�������� ����.'���½�û�����ο� ����(����ڰ� ���� ��츸)

	dim WaitExist : WaitExist = false
	sqlStr = " select count(*) as cnt from " + vbCrlf
    sqlStr = sqlStr + " db_academy.dbo.tbl_academy_order_master m " + vbCrlf
    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_academy_order_detail d" + vbCrlf
    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbCrlf
    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_lec_waiting_user w" + vbCrlf
    sqlStr = sqlStr + " 	on d.itemid=w.lec_idx" + vbCrlf
    sqlStr = sqlStr + " 	and d.itemoption=w.lecOption" + vbCrlf
    sqlStr = sqlStr + " 	and w.isusing='Y'" + vbCrlf
    sqlStr = sqlStr + " 	and w.currstate<7" + vbCrlf
    sqlStr = sqlStr + " 	and IsNULL(w.regendday,'9999-12-12')>getdate()" + vbCrlf
    sqlStr = sqlStr + " where m.orderserial='" + CStr(orderserial) + "'" + vbCrlf
    rsAcademyget.Open sqlStr,dbAcademyget,1
    if Not rsAcademyget.Eof then
    	WaitExist = (rsAcademyget("cnt")>0)
    end if
    rsAcademyget.Close


	if (Not WaitExist) then
    	sqlStr = "update [db_academy].[dbo].tbl_lec_item_option " + vbCrlf
    	sqlStr = sqlStr + " set limit_sold=limit_sold - T.cnt" + vbCrlf
    	sqlStr = sqlStr + " from " + vbCrlf
    	sqlStr = sqlStr + " (select d.itemid, d.itemoption, sum(d.itemno) as cnt" + vbCrlf
    	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d" + vbCrlf
    	sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
    	sqlStr = sqlStr + " and d.itemid<>0" + vbCrlf
    	sqlStr = sqlStr + " group by d.itemid, d.itemoption ) as T" + vbCrlf
    	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item_option.lecidx=T.Itemid"
    	sqlStr = sqlStr + " and [db_academy].[dbo].tbl_lec_item_option.lecoption=T.itemoption"
    	rsAcademyget.Open sqlStr,dbAcademyget,1

    	sqlStr = "update [db_academy].[dbo].tbl_lec_item" + vbCrlf
        sqlStr = sqlStr + " set limit_count=T.limit_count" + vbCrlf
        sqlStr = sqlStr + " ,limit_sold=T.limit_sold" + vbCrlf
        sqlStr = sqlStr + " ,wait_count=T.wait_count" + vbCrlf
        sqlStr = sqlStr + " from (" + vbCrlf
        sqlStr = sqlStr + " 	select o.lecidx, sum(limit_count) as limit_count, sum(limit_sold) as limit_sold" + vbCrlf
        sqlStr = sqlStr + " 	,sum(wait_count) as wait_count" + vbCrlf
        sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_lec_item_option o" + vbCrlf
        sqlStr = sqlStr + " 		Join (select distinct itemid from [db_academy].[dbo].tbl_academy_order_detail where orderserial='" + CStr(orderserial) + "') A" + vbCrlf
        sqlStr = sqlStr + " 		on o.lecidx=A.itemid" + vbCrlf
        sqlStr = sqlStr + " 	group by o.lecidx" + vbCrlf
        sqlStr = sqlStr + " ) T" + vbCrlf
        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.lecidx" + vbCrlf

    	rsAcademyget.Open sqlStr,dbAcademyget,1
	end if
	'���½�û�����ο� ����(����ڰ� ���� ��츸)
'	if (olecture.FOneItem.FWaitCount = 0) then
'    	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
'    	sqlStr = sqlStr + " set limit_sold = limit_sold - " + CStr(oordermaster.FOneItem.Ftotalitemno) + " "
'    	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
'    	rsAcademyget.Open sqlStr,dbAcademyget,1
'    end if

    '��ȸ���� �ƴѰ��(���ϸ���/���� ó��)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Fmiletotalprice <> 0) then
        	'��븶�ϸ��� ���
        	sqlStr = " update [db_user].[dbo].tbl_mileagelog "
        	sqlStr = sqlStr + " set deleteyn='Y' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('���ϸ��� ����� ��ҵǾ����ϴ�.');</script>"
        end if

    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '������� ���밡���ϰ� ��ȯ
        	sqlStr = " update [db_user].[dbo].tbl_user_coupon "
        	sqlStr = sqlStr + " set isusing='N' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('���� ����� ��ҵǾ����ϴ�.');</script>"
        end if

        '==============================================================
        '�� ��븶�ϸ���/ȹ�渶�ϸ��� ����
        updateUserMileage oordermaster.FOneItem.FUserID
    end if

    response.write "<script>alert('���½�û�� ��ҵǾ����ϴ�.');</script>"
    if (returnmethod = "bank") then
        insertRepayBank orderserial, id, refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle
        response.write "<script>alert('������ȯ�� CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "creditcard") then
        insertCancelCardRequest orderserial, id, refundrequire, refundcstitle
        'cancelInicisCardPay orderserial
        response.write "<script>alert('ī����� CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "realtimetransfer") then
        insertCancelRealTimeTransferRequest orderserial, id, refundrequire, refundcstitle
        response.write "<script>alert('�ǽð���ü��� CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "point") then
        insertCancelPointRequest orderserial, id, refundrequire
        response.write "<script>alert('����Ʈ���� ��ҿ�û CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mall") then
        insertCancelMallRequest orderserial, id, refundrequire
        response.write "<script>alert('�ܺθ����� ��ҿ�û CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "allatcard") then
        insertCancelAllAtCardRequest orderserial, id, refundrequire
        response.write "<script>alert('�þ�ī����� ��ҿ�û CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "ticket") then
        insertCancelTicketRequest orderserial, id, refundrequire
        response.write "<script>alert('��ǰ�ǰ��� ��ҿ�û CS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mileage") then
        response.write "<script>alert('�� ����ݾ��� ���ϸ��� ��ȯ�Ǿ����ϴ�. ���� �۾���.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    else
        if ((oordermaster.FOneItem.Fsubtotalprice > 0) and (oordermaster.FOneItem.Fipkumdiv >= 4)) then
            response.write "<script>alert('ȯ�ҹ���� ���õ��� �ʾҽ��ϴ�.'); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        else
            response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        end if
    end if
end if

if (mode = "cancelitem") then
    '�κ����
    ' - �Ա��� ����� ���, �κ����AS��Ϲ�����ó��, �ֹ����̺��� ��ǰ��������, �ֹ����̺��� ȹ�渶�ϸ��� ����, ����ڸ��ϸ��� ��� ����, ����������ϸ�������, ���� �������� ������
    ' - �Ա��� ����� ���, �κ����AS��Ϲ�����ó��, �ֹ����̺��� ��ǰ��������, �ֹ����̺��� ȹ�渶�ϸ��� ����, ����ڸ��ϸ��� ��� ����, ����������ϸ�������, ���� �������� ������, �ſ�ī����� �Ǵ� �������Ա�AS����������� �̵�

    '======================================================================
    '�κ����AS��Ϲ�����ó��(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	'response.write sqlStr
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
	''rsget("refundrequire")  = 0
	''rsget("cause")          = cause
	''rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = rebankname
	'rsget("rebankaccount")  = rebankaccount
	'rsget("rebankownername")        = rebankownername
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close




    '======================================================================
    '�κ����AS��Ϲ�����ó��(��ǰ���)
    '��û������ �ش� ���������
    'TODO : �����Ͽ��� �ϳ��� ���¸� �ִٰ� �����Ѵ�.
    dmasterid = id
    dorderserial = orderserial

    contents_finish = ""
    canceldetailno = 0
    canceldetailnamelist = ""
    detailitemlist = split(detailitemlist, "|")
    detailitemnolist = split(detailitemnolist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
            for j = 0 to oorderdetail.FResultCount - 1
                if (CLng(oorderdetail.FItemList(j).Fdetailidx) = CLng(trim(detailitemlist(i)))) then
                    canceldetailno = canceldetailno + 1

                    canceldetailall = True
					sqlStr = " select itemno from " + vbCrlf
				    sqlStr = sqlStr + " db_academy.dbo.tbl_academy_order_master m " + vbCrlf
				    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_academy_order_detail d" + vbCrlf
				    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbCrlf
				    sqlStr = sqlStr + " where d.detailidx='" + CStr(detailitemlist(i)) + "'" + vbCrlf
				    rsAcademyget.Open sqlStr,dbAcademyget,1
				    if Not rsAcademyget.Eof then
				    	canceldetailall = (rsAcademyget("itemno") <= CLng(detailitemnolist(i)))
				    end if
				    rsAcademyget.Close


					if (sitename = "academy") then
	                    '����
	                    if (canceldetailnamelist = "") then
	                        canceldetailnamelist = oorderdetail.FItemList(j).Fentryname
	                    else
	                        canceldetailnamelist = canceldetailnamelist + "," + oorderdetail.FItemList(j).Fentryname
	                    end if
					else
	                    if (canceldetailnamelist = "") then
	                        canceldetailnamelist = oorderdetail.FItemList(j).FItemName & "[" & CStr(oorderdetail.FItemList(j).Fitemoptionname) & "] " & detailitemnolist(i) & " ��"
	                    else
	                        canceldetailnamelist = canceldetailnamelist + "<br>" + oorderdetail.FItemList(j).FItemName & "[" & CStr(oorderdetail.FItemList(j).Fitemoptionname) & "] " & detailitemnolist(i) & " ��"
	                    end if
					end if

					if (canceldetailall = true) then
	                    sqlStr = " update [db_academy].[dbo].tbl_academy_order_detail set cancelyn = 'Y' "
	                    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and detailidx = " + CStr(trim(detailitemlist(i))) + " "
	                    rsAcademyget.Open sqlStr,dbAcademyget,1
	                    'response.write sqlStr
					else
	                    sqlStr = " update [db_academy].[dbo].tbl_academy_order_detail set itemno = itemno - " & detailitemnolist(i) & " "
	                    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and detailidx = " + CStr(trim(detailitemlist(i))) + " "
	                    rsAcademyget.Open sqlStr,dbAcademyget,1
					end if

                    if (sitename = "academy") and (canceldetailno = 1) then
    			        sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,gubun01,gubun02) "
    			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oordermaster.FOneItem.Fitemid) + ",'" + CStr(oordermaster.FOneItem.Fitemoption) + "','" + CStr(oordermaster.FOneItem.Fmakerid) + "','" + html2db(oordermaster.FOneItem.FItemName) + "','" + html2db(oordermaster.FOneItem.FItemoptionName) + "'," + CStr(oordermaster.FOneItem.Ftotalitemno) + ",1," + CStr(oordermaster.FOneItem.Fitemcost) + ",'N','" + CStr(oordermaster.FOneItem.Fipkumdiv) + "','','') "
    			        rsget.Open sqlStr,dbget,1
    			        'response.write sqlStr
    			    end if

					if (sitename <> "academy") then
				        sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,gubun01,gubun02) "
				        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(detailitemnolist(i)) + "," & detailitemnolist(i) & "," + CStr(oorderdetail.FItemList(j).Freducedprice) + ",'Y','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','','') "
				        rsget.Open sqlStr,dbget,1
				        'response.write sqlStr
			    	end if
                end if
            next
		end if
	next

    if (sitename = "academy") and (canceldetailno > 1) then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_detail set confirmitemno = " + CStr(canceldetailno) + " "
        sqlStr = sqlStr + " where masterid = " + CStr(dmasterid) + " "
        rsget.Open sqlStr,dbget,1
        'response.write sqlStr
    end if

	if (sitename = "academy") then
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
		sqlStr = sqlStr + " set finishdate=getdate() "
		sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
		sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + CStr(canceldetailno) + " ��[" + CStr(html2db(canceldetailnamelist)) + "]" + ")") + "' "
		sqlStr = sqlStr + " ,currstate = 'B007' "
		'sqlStr = sqlStr + " ,refundresult = 0 "
		sqlStr = sqlStr + " where id=" + CStr(dmasterid) + " "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	else
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
		sqlStr = sqlStr + " set finishdate=getdate() "
		sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
		sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(DIY��ǰ)") + "' "
		sqlStr = sqlStr + " ,currstate = 'B007' "
		'sqlStr = sqlStr + " ,refundresult = 0 "
		sqlStr = sqlStr + " where id=" + CStr(dmasterid) + " "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	'======================================================================
	'���½�û�����ο� ����
	if (sitename = "academy") and (olecture.FOneItem.FWaitCount = 0) then
    	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
    	sqlStr = sqlStr + " set limit_sold = limit_sold - " + CStr(canceldetailno) + " "
    	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
    	rsAcademyget.Open sqlStr,dbAcademyget,1
    end if

	'======================================================================
	'�������������̺� ���� ������Ʈ
    recalculateOrderMaster orderserial

	'======================================================================
	'�� ��븶�ϸ���/ȹ�渶�ϸ��� ����
    if (oordermaster.FOneItem.FUserID <> "") then
        updateUserMileage oordermaster.FOneItem.FUserID
    end if


    response.write "<script>alert('�κ���ҽ�û�� ��ϵǾ����ϴ�.');</script>"
    if (returnmethod = "bank") then
        insertRepayBank orderserial, dmasterid, refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle
        response.write "<script>alert('������ȯ�ҿ�ûCS �� ��ϵǾ����ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mileage") then
        response.write "<script>alert('�� ����ݾ��� ���ϸ��� ��ȯ�Ǿ����ϴ�. ���� �۾���.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    else
        if ((oordermaster.FOneItem.Fsubtotalprice > 0) and (oordermaster.FOneItem.Fipkumdiv >= 4)) then
            response.write "<script>alert('ȯ�ҹ���� ���õ��� �ʾҽ��ϴ�.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        else
            response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        end if
    end if
end if

if (mode = "revalidate") then
    '������ȯ
    ' - �������ֹ��� �ֹ����������̸�, ��ҵ� �ֹ��� ������ȯ �մϴ�.

    '======================================================================
    '������ȯAS��Ϲ�����ó��(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
'	rsget("refundrequire")  = 0
	'rsget("cause")          = cause
	'rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
'	rsget("rebankname")     = rebankname
'	rsget("rebankaccount")  = rebankaccount
'	rsget("rebankownername")        = rebankownername
'	rsget("refundbeasongpay")       = 0
'	rsget("refunditemcostsum")      = 0
'	rsget("refunddeliverypay")      = 0
'	rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close

	sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " set finishdate=getdate() "
	sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	sqlStr = sqlStr + " ,contents_finish = '" + html2db("���½�û������ȯ") + "' "
	sqlStr = sqlStr + " ,currstate = 'B007' "
	''sqlStr = sqlStr + " ,refundresult = 0 "
	sqlStr = sqlStr + " where id=" + CStr(id) + " "
	rsget.Open sqlStr,dbget,1

	'======================================================================
	'�ֹ�������ȭ
	sqlStr = " update [db_academy].[dbo].tbl_academy_order_master "
	sqlStr = sqlStr + " set cancelyn='N' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
	rsAcademyget.Open sqlStr,dbAcademyget,1

	'======================================================================
	'���½�û�����ο� ����
	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
	sqlStr = sqlStr + " set limit_sold = limit_sold + " + CStr(oordermaster.FOneItem.Ftotalitemno) + " "
	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
	rsAcademyget.Open sqlStr,dbAcademyget,1

    '��ȸ���� �ƴѰ��(���ϸ���/���� ó��)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Fmiletotalprice <> 0) then
        	'��븶�ϸ��� ����ȭ
        	sqlStr = " update [db_user].[dbo].tbl_mileagelog "
        	sqlStr = sqlStr + " set deleteyn='N' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('���ϸ��� ����� ����ȭ�Ǿ����ϴ�.');</script>"
        end if

    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '������� ����ȭ
        	sqlStr = " update [db_user].[dbo].tbl_user_coupon "
        	sqlStr = sqlStr + " set isusing='Y' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('���� ����� ����ȭ�Ǿ����ϴ�.');</script>"
        end if

        '==============================================================
        '�� ��븶�ϸ���/ȹ�渶�ϸ��� ����
        updateUserMileage oordermaster.FOneItem.FUserID
    end if

    response.write "<script>alert('���½�û�� ����ȭ�Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
end if


if (mode = "receveupche") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End

        '��ü��ǰ
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���, ��������(�κ���ҺҰ�)
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���,
        ' - ���õ� ��ǰ����� �����ϰ�, ��ǰ/ȸ�� CS �� �������·� �����Ѵ�.
        ' - ����, ��ǰ�� ��������, �������·� ����� CS �� �Ϸ�ó���ϸ鼭, ���̳ʽ� �ֹ��� ������ȯ�� CS ����� �Ѵ�. �ο��� ���ϸ��� ȸ���� ���̳ʽ� �ֹ����� ó���Ѵ�.(�ٸ� ��ƾ���� ó���Ѵ�.)

        '======================================================================
        '��ü��ǰAS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "Y"
	rsget("makerid")        = makerid
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = rebankname
	rsget("rebankaccount")  = html2db(rebankaccount)
	rsget("rebankownername")        = html2db(rebankownername)
	rsget("refundbeasongpay")       = refundbeasongpay
	rsget("refunditemcostsum")      = refunditemcostsum
	rsget("refunddeliverypay")      = refunddeliverypay
	rsget("refundadjustpay")        = refundadjustpay

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '��ü��ǰAS���(��ǰ���)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '���õ� ��ǰ�� ���Ϸᰡ �ƴ� ��ǰ�� �ִ��� üũ
                                'if (oorderdetail.FItemList(i).GetStateName <> "���Ϸ�") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('��ǰ�� ������ ���� ��ǰ�� �ֽ��ϴ�. ����� ��ҵ˴ϴ�.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('��ǰ������ ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "recevetenten") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        'ȸ����û
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���, ��������(�κ���ҺҰ�)
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���,
        ' - ���õ� ��ǰ����� �����ϰ�, ��ǰ/ȸ�� CS �� �������·� �����Ѵ�.
        ' - ����, ��ǰ�� ��������, �������·� ����� CS �� �Ϸ�ó���ϸ鼭, ���̳ʽ� �ֹ��� ������ȯ�� CS ����� �Ѵ�. �ο��� ���ϸ��� ȸ���� ���̳ʽ� �ֹ����� ó���Ѵ�.(�ٸ� ��ƾ���� ó���Ѵ�.)

        '======================================================================
        'ȸ����ûAS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = rebankname
	rsget("rebankaccount")  = html2db(rebankaccount)
	rsget("rebankownername")        = html2db(rebankownername)
	rsget("refundbeasongpay")       = refundbeasongpay
	rsget("refunditemcostsum")      = refunditemcostsum
	rsget("refunddeliverypay")      = refunddeliverypay
	rsget("refundadjustpay")        = refundadjustpay

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        'ȸ����ûAS���(��ǰ���)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '���õ� ��ǰ�� ���Ϸᰡ �ƴ� ��ǰ�� �ִ��� üũ
                                'if (oorderdetail.FItemList(i).GetStateName <> "���Ϸ�") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('��ǰ�� ������ ���� ��ǰ�� �ֽ��ϴ�. ����� ��ҵ˴ϴ�.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('ȸ����û�� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "change") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '�±�ȯ
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���, ��������(�κ���ҺҰ�)
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���,
        ' - ���õ� ��ǰ����� �����ϰ�, �±�ȯ CS �� �������·� �����Ѵ�.
        ' - ����, ���� ������ �����鼭, �����ȣ�� �Է��ϰ� ����ó���Ѵ�.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '�±�ȯAS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '�±�ȯAS���(��ǰ���)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '���õ� ��ǰ�� ���Ϸᰡ �ƴ� ��ǰ�� �ִ��� üũ
                                'if (oorderdetail.FItemList(i).GetStateName <> "���Ϸ�") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('��ǰ�� ������ ���� ��ǰ�� �ֽ��ϴ�. ����� ��ҵ˴ϴ�.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('�±�ȯ�� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "omit") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '������߼�
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���, ��������(�κ���ҺҰ�)
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���,
        ' - ���õ� ��ǰ����� �����ϰ�, ������߼� CS �� �������·� �����Ѵ�.
        ' - ����, ���� ������ �����鼭, �����ȣ�� �Է��ϰ� ����ó���Ѵ�.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '������߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '������߼�AS���(��ǰ���)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '���õ� ��ǰ�� ���Ϸᰡ �ƴ� ��ǰ�� �ִ��� üũ
                                'if (oorderdetail.FItemList(i).GetStateName <> "���Ϸ�") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('��ǰ�� ������ ���� ��ǰ�� �ֽ��ϴ�. ����� ��ҵ˴ϴ�.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('������߼��� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "more") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '���񽺹߼�
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���, ��������(�κ���ҺҰ�)
        '���õ� ��ǰ��, ���Ϸᰡ �ƴ� ��ǰ�� ���� ���,
        ' - ���õ� ��ǰ����� �����ϰ�, ���񽺹߼� CS �� �������·� �����Ѵ�.
        ' - ����, ���� ������ �����鼭, �����ȣ�� �Է��ϰ� ����ó���Ѵ�.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("�κ����(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '���񽺹߼�AS���(��ǰ���)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '���õ� ��ǰ�� ���Ϸᰡ �ƴ� ��ǰ�� �ִ��� üũ
                                'if (oorderdetail.FItemList(i).GetStateName <> "���Ϸ�") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('��ǰ�� ������ ���� ��ǰ�� �ֽ��ϴ�. ����� ��ҵ˴ϴ�.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('���񽺹߼��� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelcard") then
        '�ſ�ī��/��ǰ��/�ǽð���ü��ҿ�û

        '======================================================================
        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('�ſ�ī��/��ǰ��/�ǽð���ü��� ��û�� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelbank") then
        'ȯ�ҿ�û

        '======================================================================
        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('ȯ�ҿ�û�� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelothersite") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '�ܺθ���ҿ�û

        '======================================================================
        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('�ܺθ���ҿ�û�� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "writereadme") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '������ǻ���

        '======================================================================
        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('������ǻ����� ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "writeetcnote") then
    response.write "������ ���� ���"
    dbget.close()	:	response.End
        '��Ÿ����

        '======================================================================
        '���񽺹߼�AS���(����Ÿ)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('��Ÿ������ ��ϵǾ����ϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if



























sub recalculateOrderMaster(byVal orderserial)
	dim sqlStr
	dim jumundiv, discountrate, linkorderserial, miletotalprice, tencardspend, spendmembership, userid, sitename, ipkumdiv
	dim itemcostsum, itemvatsum, itemmileagesum, deliverpay, minusitemcostsum
	dim discountitemcostsum, discountitemvatsum, discountminusitemcostsum
	dim isallreturn, hasreturn, notreturnsubtotal
	dim subtotal, totalsum, totalitemno, cancelitemno, cancelprice

	sqlStr = " select * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            jumundiv = rsAcademyget("jumundiv")
            discountrate = rsAcademyget("discountrate")
            linkorderserial = rsAcademyget("linkorderserial")
            miletotalprice = rsAcademyget("miletotalprice")
            tencardspend = rsAcademyget("tencardspend")
            spendmembership = rsAcademyget("spendmembership")

            userid = rsAcademyget("userid")
            sitename = rsAcademyget("sitename")
            ipkumdiv = rsAcademyget("ipkumdiv")
    else
            jumundiv = "0"
            discountrate = 1.0
            linkorderserial = ""
            miletotalprice = 0
            tencardspend = 0
            spendmembership = 0

            userid = ""
            sitename = ""
            ipkumdiv = "0"
    end if
    rsAcademyget.close

	'�����հ� ���ϱ�
    sqlStr = "          select   sum((case when cancelyn = 'Y' then 0 else itemcost end) * itemno) as itemcostsum "
    sqlStr = sqlStr + "         ,sum((case when cancelyn = 'Y' then 0 else mileage end) * itemno) as itemmileagesum "
    sqlStr = sqlStr + "         ,sum((case when ((cancelyn <> 'Y') and (itemno < 0)) then itemcost else 0 end) * itemno) as minusitemcostsum "
    sqlStr = sqlStr + "         ,sum((case when cancelyn = 'Y' then 0 else round((" + CStr(discountrate) + " * itemcost), 2) end) * itemno) as discountitemcostsum "
    sqlStr = sqlStr + "         ,sum((case when ((cancelyn <> 'Y') and (itemno < 0)) then round((" + CStr(discountrate) + " * itemcost), 2) else 0 end) * itemno) as discountminusitemcostsum "
    sqlStr = sqlStr + "         ,sum(case when cancelyn <> 'Y' then itemno else 0 end) as totalitemno "
    sqlStr = sqlStr + "         ,sum(case when cancelyn = 'Y' then itemno else 0 end) as cancelitemno "
    sqlStr = sqlStr + "         ,sum((case when cancelyn <> 'Y' then 0 else itemcost end) * itemno) as cancelprice "
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail "
    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1
    'response.write sqlStr

    if Not rsAcademyget.Eof then
        itemcostsum = rsAcademyget("itemcostsum")
        itemmileagesum = rsAcademyget("itemmileagesum")
        deliverpay = 0
        minusitemcostsum = rsAcademyget("minusitemcostsum")

        discountitemcostsum = rsAcademyget("discountitemcostsum")
        discountitemvatsum = 0
        discountminusitemcostsum = rsAcademyget("discountminusitemcostsum")

        totalitemno = rsAcademyget("totalitemno")
        cancelitemno = rsAcademyget("cancelitemno")
        cancelprice = rsAcademyget("cancelprice")
    else
        itemcostsum = 0
        itemmileagesum = 0
        deliverpay = 0
        minusitemcostsum = 0

        discountitemcostsum = 0
        discountitemvatsum = 0
        discountminusitemcostsum = 0

        totalitemno = 0
        cancelitemno = 0
        cancelprice = 0
    end if
    rsAcademyget.close

    '��ü��ǰ/�κй�ǰ Ȯ��
    if (linkorderserial<>"") and (jumundiv="9") then
        if (discountminusitemcostsum < 0) then
                hasreturn = "Y"
        end if

        if (discountitemcostsum = discountminusitemcostsum) then
                isallreturn = "Y"
        end if

        notreturnsubtotal = discountitemcostsum - discountminusitemcostsum
    end if

    subtotal = discountitemcostsum + deliverpay

	if (jumundiv<>"9") then
		subtotal = subtotal - miletotalprice - tencardspend - spendmembership
	else
		if (isallreturn = "Y") or (Abs(miletotalprice + tencardspend + spendmembership) > Abs(notreturnsubtotal)) then
			'��ü��ǰ�ΰ��
			'�κй�ǰ�ΰ�� : (�����űݾ�-��ǰ�ݾ�)�� (����+���ϸ������)�ݾ׺��� �������
			subtotal = subtotal + miletotalprice + tencardspend + spendmembership
		end if
	end if

    totalsum = itemcostsum + deliverpay

    sqlStr = "update [db_academy].[dbo].tbl_academy_order_master set " + vbCrlf
	'sqlStr = sqlStr & " totalvat=" & itemvatsum & "," + vbCrlf
	sqlStr = sqlStr & " totalitemno=" & totalitemno & "," + vbCrlf
	sqlStr = sqlStr & " cancelitemno=" & cancelitemno & "," + vbCrlf
	sqlStr = sqlStr & " cancelprice=" & cancelprice & "," + vbCrlf
	'sqlStr = sqlStr & " totalcost=" & totalsum & "," + vbCrlf
	sqlStr = sqlStr & " totalsum=" & totalsum & "," + vbCrlf
	sqlStr = sqlStr & " totalmileage=" & itemmileagesum & "," + vbCrlf
	sqlStr = sqlStr & " subtotalprice=" & subtotal & vbCrlf
	sqlStr = sqlStr & " where orderserial='" + CStr(orderserial) + "' "
	'response.write sqlStr
	rsAcademyget.Open sqlStr,dbAcademyget,1


	'if (userid<>"") and ((sitename="10x10") or (sitename="way2way")) and (CInt(ipkumdiv)>3) then
	'	sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + vbCrlf
	'	sqlStr = sqlStr + " set [db_user].[dbo].tbl_user_current_mileage.jumunmileage=T.totmile" + vbCrlf
	'	sqlStr = sqlStr + " from " + vbCrlf
	'	sqlStr = sqlStr + " (select sum(totalmileage) as totmile" + vbCrlf
	'	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master" + vbCrlf
	'	sqlStr = sqlStr + " where userid='" + userid + "'" + vbCrlf
	'	sqlStr = sqlStr + " and sitename in ('10x10','way2way')" + vbCrlf
	'	sqlStr = sqlStr + " and cancelyn='N'" + vbCrlf
	'	sqlStr = sqlStr + " and ipkumdiv>3" + vbCrlf
	'	sqlStr = sqlStr + " ) as T" + vbCrlf
	'	sqlStr = sqlStr + " where [db_user].[dbo].tbl_user_current_mileage.userid='" + userid + "'"
	'	rsget.Open sqlStr,dbget,1
	'end if
end sub

sub updateUserMileage(byVal userid)
	dim sqlStr

	'// ���ʽ�/��븶�ϸ��� ��� ����(�ű�Proc)
	sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
	dbget.Execute sqlStr

	dim totmile
	sqlStr = " select IsNULL(sum(totalmileage),0) as totmile"
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    sqlStr = sqlStr + " and cancelyn='N'"
    sqlStr = sqlStr + " and ipkumdiv>3"
    rsAcademyget.Open sqlStr,dbAcademyget,1
    if Not rsAcademyget.Eof then
    	totmile = rsAcademyget("totmile")
    else
    	totmile = 0
    end if
    rsAcademyget.Close


	'==============================================================
	'�ֹ����ϸ��� ��� ����([db_academy].[dbo].tbl_academy_order_master)
    sqlStr = "update [db_user].[dbo].tbl_user_current_mileage"
    sqlStr = sqlStr + " set academymileage=" + CStr(totmile) + ""
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    rsget.Open sqlStr,dbget,1
end sub

sub insertRepayBank(byVal orderserial, byVal basecsid, byVal refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid
    dim id
    dim orgsubtotalprice

'    sqlStr = " select top 1 * from [db_cs].[dbo].tbl_new_as_list where id = " + CStr(basecsid) + " "
'    rsget.Open sqlStr,dbget,1
'
'    if Not rsget.Eof then
'            rebankname = db2html(rsget("rebankname"))
'            rebankaccount = db2html(rsget("rebankaccount"))
'            rebankownername = db2html(rsget("rebankownername"))
'
'            cause = ""
'            causedetail = ""
'    else
'            rebankname = ""
'            rebankaccount = ""
'            rebankownername = ""
'            refundrequire = 0
'
'            cause = ""
'            causedetail = ""
'    end if
'    rsget.close

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            orgsubtotalprice = "0"
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A003"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle)
	rsget("contents_jupsu") = html2db("���¹�ȣ : " + rebankaccount + " / ���� : " + rebankname + " / ������ : " + rebankownername + " ")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	'rsget("requireupche")   = "N"
	'rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = html2db(rebankname)
	'rsget("rebankaccount")  = html2db(rebankaccount)
	'rsget("rebankownername")        = html2db(rebankownername)

	rsget.update
	id = rsget("id")
	rsget.close


	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R007" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,'" & rebankname &"'"
    sqlStr = sqlStr + " ,'" & rebankaccount &"'"
    sqlStr = sqlStr + " ,'" & rebankownername &"'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'�ſ�ī�����
sub insertCancelCardRequest(byVal orderserial, byVal basecsid, byVal refundrequire, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid, paygatetid
    dim id
    dim orgsubtotalprice

    cause = ""
    causedetail = ""

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = db2html(rsAcademyget("paygatetid"))
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            orgsubtotalprice = 0
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A007"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle + " - ī�����")
	rsget("contents_jupsu") = html2db("TID[ " + paygatetid + " ]")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close


	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername, paygateTid"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R100" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,'" & paygatetid & "'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'�ǽð���ü ���
sub insertCancelRealTimeTransferRequest(byVal orderserial, byVal basecsid, byVal refundrequire, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid, paygatetid
    dim id
    dim orgsubtotalprice

    cause = ""
    causedetail = ""

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = db2html(rsAcademyget("paygatetid"))
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            orgsubtotalprice = 0
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A007"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle + " - �ǽð���ü���")
	rsget("contents_jupsu") = html2db("�ǽð���ü[ " + paygatetid + " ]")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close



	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername, paygateTid"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R020" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,'" & paygatetid & "'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'����Ʈ ���
sub insertCancelPointRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("�������(����Ʈ���)")
	rsget("contents_jupsu") = html2db("����Ʈ[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'������ ���
sub insertCancelMallRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("�������(���������)")
	rsget("contents_jupsu") = html2db("������[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'�þ�ī�� ���
sub insertCancelAllAtCardRequest(byVal orderserial, byVal basecsid, byVal refundrequire)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("�������(�þ�ī�����)")
	rsget("contents_jupsu") = html2db("�þ�ī��[ " + paygatetid + " ]")
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'��ǰ�� ���
sub insertCancelTicketRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("�������(��ǰ�����)")
	rsget("contents_jupsu") = html2db("��ǰ��[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

sub cancelInicisCardPay(byVal orderserial)
    dim sqlStr
    dim refundrequire, cause, causedetail
    dim buyname, userid, paygatetid, accountdiv
    dim id

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1
    'response.write sqlStr

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = rsAcademyget("paygatetid")
            accountdiv = rsAcademyget("accountdiv")
            if ((accountdiv <> "20") and (accountdiv <> "90") and (accountdiv <> "100")) then
                    paygatetid = ""
            end if
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            accountdiv = ""
    end if
    rsAcademyget.close

    '�̴Ͻý� ���(ī�����)
    dim INIpay, PInst, ResultCode, ResultMsg

    ResultCode = "--"
    ResultMsg = "TID ����"
    if (paygatetid <> "") then
        Set INIpay = Server.CreateObject("INItx41.INItx41.1")
        PInst = INIpay.Initialize("")
        INIpay.SetActionType CLng(PInst), "CANCEL"

        INIpay.SetField CLng(PInst), "pgid", "IniTechPG_"       'PG ID (����)
        INIpay.SetField CLng(PInst), "spgip", "203.238.3.10"    '���� PG IP (����)
        INIpay.SetField CLng(PInst), "mid", "teenxteen3"        '�������̵�
        INIpay.SetField CLng(PInst), "admin", "1111"            'Ű�н�����(�������̵� ���� ����)
        INIpay.SetField CLng(PInst), "tid", paygatetid          '����� �ŷ���ȣ(TID)
        INIpay.SetField CLng(PInst), "msg", "CSī�����"        '��� ����
        INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
        INIpay.SetField CLng(PInst), "debug", "true"            '�α׸��("true"�� �����ϸ� ���� �α׸� ����)
        INIpay.SetField CLng(PInst), "merchantreserved", "����" '����

        INIpay.StartAction(CLng(PInst))

        ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ��Ҽ���)
        ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
        'CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '�̴Ͻý� ��ҳ�¥
        'CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '�̴Ͻý� ��ҽð�
        'Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '���ݿ����� ��� ���ι�ȣ

        INIpay.Destroy CLng(PInst)
    end if

    'response.write ResultMsg
end sub

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
