<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.08 �ѿ�� ����
'###########################################################

function CsState2Name_off(byval v)
	if IsNull(v) or (v="") then
		Exit function
	end if

	if v="0" then

	elseif v="B001" then
		CsState2Name_off = "����"
	elseif v="B004" then
		CsState2Name_off = "������Է�"
	elseif v="B003" then
		CsState2Name_off = ""
	elseif v="B006" then
		CsState2Name_off = "��üó���Ϸ�"
	elseif v="B007" then
		CsState2Name_off = "ó���Ϸ�"
	elseif v="B008" then
		CsState2Name_off = "����ó���Ϸ�"		
	else
	end if
end function

function AddOneCSDetail_off(csmasteridx, dorderdetailidx, orderno, dregitemno)
    dim sqlStr , jumundetailidx , jumunitemgubun

	if masteridx = "" then exit function    

	'/���� �Ǹ����̺� �� detailidx
	sqlStr = "select "
	sqlStr = sqlStr & " detailidx ,masteridx ,orgdetailidx , itemno,itemgubun"
	sqlStr = sqlStr & " from [db_shop].dbo.tbl_shopbeasong_order_detail"
	sqlStr = sqlStr & " where masteridx = "&masteridx&""

    'response.write sqlStr &"<Br>"
    rsget.Open sqlStr,dbget,1    
	    if Not rsget.Eof then
			jumundetailidx = rsget("orgdetailidx")
			jumunitemgubun = rsget("itemgubun")
	    end if    
    rsget.Close

	sqlStr = ""
    sqlStr = " insert into [db_shop].dbo.tbl_shopbeasong_cs_detail"
    sqlStr = sqlStr + " (masteridx, orderdetailidx,orderno, itemid, itemoption,makerid"
    sqlStr = sqlStr + " , regitemno, confirmitemno,orderitemno,jumundetailidx ,itemgubun) "
    sqlStr = sqlStr + " values(" + CStr(csmasteridx) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(orderno) + "'"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,"&jumundetailidx&""
    sqlStr = sqlStr + " ,'"&jumunitemgubun&"'"    
    sqlStr = sqlStr + " )"
    
	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

function RegCSMasterAddUpche_off(csmasteridx, imakerid)
    dim sqlStr
    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master"    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requireupche='Y'"               + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

function RegCSMasterAddmaejang_off(csmasteridx)
    dim sqlStr
    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master"    + VbCrlf
    sqlStr = sqlStr + " set requiremaejang='Y'"+ VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

''�ٷ� �Ϸ� ó���� ���� ���� ����.
function IsDirectProceedFinish_off(divcd, csmasteridx, masteridx, byRef EtcStr)
    dim sqlStr
    dim cancelyn, ipkumdiv
    IsDirectProceedFinish_off = false

    if (divcd="A008") then
        '' ��ϵ� ��ǰ�� ��ǰ�غ��� ���°� ������ �������·� ����
        sqlStr = " select count(*) as invalidcount"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d"
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_cs_detail c"
		sqlStr = sqlStr + " 	on d.detailidx = c.orderdetailidx "    
        sqlStr = sqlStr + " where d.itemid<>0"
		sqlStr = sqlStr + " and c.masteridx='" + CStr(csmasteridx) + "'"
        sqlStr = sqlStr + " and m.masteridx='" + masteridx + "'"
        sqlStr = sqlStr + " and d.currstate>=3"
        sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		
	    'response.write sqlStr &"<Br>"
	    rsget.Open sqlStr,dbget,1    
	    if Not rsget.Eof then
        	IsDirectProceedFinish_off = (rsget("invalidcount")=0)
        end if
        rsget.close
    else
    end if
end function

'/�ֹ���� ������
function CancelProcess_off(byval detailitemlist, csmasteridx, orderno,masteridx,cancelorderno)
    dim sqlStr, result ,tmp, buf ,i ,dorderdetailidx, dregitemno

    'if cancelorderno = "" then exit function
    if detailitemlist = "" and csmasteridx = "" then exit function
	
	if detailitemlist <> "" then
	    buf = split(detailitemlist, "|")
	    
	    for i = 0 to UBound(buf)
			if (TRIM(buf(i)) <> "") then
				tmp = split(buf(i), Chr(9))
	
				dorderdetailidx = tmp(0)
				dregitemno      = tmp(1)
	
				sqlStr = ""
				sqlStr = "update d set" + vbcrlf 
				sqlStr = sqlStr & " d.cancelyn = 'Y'"	'  , d.cancelorgdetailidx = cod.idx
				sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail d"
				sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail od"
				sqlStr = sqlStr & " 	on d.orgdetailidx = od.idx"
				sqlStr = sqlStr & " 	and d.detailidx = "&dorderdetailidx&""
				sqlStr = sqlStr & "		and d.cancelyn = 'N'"
				sqlStr = sqlStr & "		and od.cancelyn = 'N'"
'				sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail cod"
'				sqlStr = sqlStr & " 	on od.itemid = cod.itemid"
'				sqlStr = sqlStr & " 	and od.itemgubun = cod.itemgubun"
'				sqlStr = sqlStr & " 	and od.itemoption = cod.itemoption"
'				sqlStr = sqlStr & " 	and cod.orderno = '"&cancelorderno&"'"
'				sqlStr = sqlStr & "		and cod.cancelyn = 'N'"
	
				'response.write sqlStr &"<Br>"
			    dbget.Execute sqlStr
			end if
		next
	else
		
		sqlStr = "update d set" + vbcrlf 
		sqlStr = sqlStr & " d.cancelyn = 'Y'"	'  , d.cancelorgdetailidx = cod.idx
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_cs_detail csd"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopbeasong_order_detail d"
		sqlStr = sqlStr & " 	on csd.orderdetailidx = d.detailidx"
		sqlStr = sqlStr & " 	and csd.masteridx = "&csmasteridx&""
		sqlStr = sqlStr & "		and d.cancelyn = 'N'"
'		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail cod"
'		sqlStr = sqlStr & " 	on csd.itemid = cod.itemid"
'		sqlStr = sqlStr & " 	and csd.itemgubun = cod.itemgubun"
'		sqlStr = sqlStr & " 	and csd.itemoption = cod.itemoption"
'		sqlStr = sqlStr & " 	and cod.orderno = '"&cancelorderno&"'"
'		sqlStr = sqlStr & "		and cod.cancelyn = 'N'"
				
		'response.write sqlStr &"<Br>"
	    dbget.Execute sqlStr
	end if
end function

'/�ֹ���� ������
function masterCancelProcess_off(masteridx,cancelorderno)
    dim sqlStr

    if cancelorderno = "" then exit function
    if masteridx = "" then exit function

	sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master set" + vbcrlf
	sqlStr = sqlStr & " cancelorgorderno='"&cancelorderno&"'" + vbcrlf
	sqlStr = sqlStr & " where masteridx = "&masteridx&""

	'response.write sqlStr &"<br>"
	dbget.execute sqlStr
end function

'/�ֹ������� ���̳ʽ��ֹ��� ��ġ�ϴ��� üũ
function GetPartialCancelRegValidResult_off(byval detailitemlist, csmasteridx, orderno,masteridx,cancelorderno)
    dim sqlStr, result ,tmp, buf ,i ,dorderdetailidx, dregitemno

    if cancelorderno = "" then exit function
    if detailitemlist = "" then exit function
    GetPartialCancelRegValidResult_off = ""
    result = ""

    buf = split(detailitemlist, "|")
    
    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dregitemno      = tmp(1)

			sqlStr = ""
			sqlStr = "select top 1000" 
			sqlStr = sqlStr & " d.detailidx ,d.itemid ,d.itemoption ,d.itemgubun"
			sqlStr = sqlStr & " ,od.sellprice as odsellprice"
			sqlStr = sqlStr & " ,cod.itemno ,cod.sellprice as codsellprice,cod.idx as codorgdetailidx"
			sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail d"
			sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail od"
			sqlStr = sqlStr & " 	on d.orgdetailidx = od.idx"
			sqlStr = sqlStr & " 	and d.detailidx = "&dorderdetailidx&""
			sqlStr = sqlStr & "		and d.cancelyn = 'N'"
			sqlStr = sqlStr & "		and od.cancelyn = 'N'"
			sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail cod"
			sqlStr = sqlStr & " 	on od.itemid = cod.itemid"
			sqlStr = sqlStr & " 	and od.itemgubun = cod.itemgubun"
			sqlStr = sqlStr & " 	and od.itemoption = cod.itemoption"
			sqlStr = sqlStr & " 	and cod.orderno = '"&cancelorderno&"'"
			sqlStr = sqlStr & "		and cod.cancelyn = 'N'"
	
		    'response.write sqlStr &"<Br>"
		    rsget.Open sqlStr,dbget,1    
		    if Not rsget.Eof then
				if rsget("odsellprice") <> rsget("codsellprice") then
					GetPartialCancelRegValidResult_off = "[��ǰ�ڵ�:" & rsget("itemid") & "]�ֹ��Ͻ� ������ ���̳ʽ� �ֹ������� �ǸŰ����� Ʋ���ϴ�"					
				end if
				if rsget("itemno") <> dregitemno*-1 then
					GetPartialCancelRegValidResult_off = "[��ǰ�ڵ�:" & rsget("itemid") & "]�ֹ��Ͻ� ������ ���̳ʽ� �ֹ������� ������ Ʋ���ϴ�"					
				end if
			else
				GetPartialCancelRegValidResult_off = "�ֹ������� ���̳ʽ� �ֹ������� ��ġ���� �ʽ��ϴ�"
		    end if    
		    rsget.Close
		end if
	next
end function

function AddCSDetailByArrStr_off(byval detailitemlist, csmasteridx, orderno,masteridx)
    dim sqlStr, tmp, buf, i ,dorderdetailidx, dregitemno    

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dregitemno      = tmp(1)

	        call AddOneCSDetail_off(csmasteridx, dorderdetailidx,orderno , dregitemno)
		end if
	next

	sqlStr = " update [db_shop].dbo.tbl_shopbeasong_cs_detail"
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , itemgubun=T.itemgubun"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , orderitemno=T.itemno"
	sqlStr = sqlStr + " , isupchebeasong=T.isupchebeasong"
	sqlStr = sqlStr + " , regdetailstate=T.currstate"
	sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopbeasong_order_detail T"
	sqlStr = sqlStr + " where T.orderno='" + orderno + "'"
	sqlStr = sqlStr + " and [db_shop].dbo.tbl_shopbeasong_cs_detail.masteridx=" + CStr(csmasteridx)
	sqlStr = sqlStr + " and [db_shop].dbo.tbl_shopbeasong_cs_detail.orderdetailidx=T.detailidx"
	
	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

'' CS Master ����
function RegCSMaster_off(divcd, orderno,reguserid, title, contents_jupsu,masteridx)
    dim sqlStr, InsertedId
	
	sqlStr = ""
    sqlStr = " select * from db_shop.dbo.tbl_shopbeasong_cs_master where 1=0 "
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    
    	rsget("orgmasteridx") = masteridx
        rsget("divcd")          = divcd
    	rsget("orderno")    = orderno
    	rsget("customername")   = ""    	
    	rsget("writeuser")      = reguserid
    	rsget("title")          = title
    	rsget("contents_jupsu") = contents_jupsu
    	rsget("currstate")      = "B001"
    	rsget("deleteyn")       = "N"

        ''''''''''''''''''''''''''''''''''
    	''rsget("requireupche")   = "N"
    	''rsget("makerid")        = ""
    	''''''''''''''''''''''''''''''''''

    rsget.update
	    InsertedId = rsget("masteridx")
	rsget.close

	dim opentitle, opencontents

	opentitle = GetDefaultTitle_off(divcd, InsertedId, orderno ,masteridx)
	
	sqlStr = ""
	sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master"  + VbCrlf
	sqlStr = sqlStr + " set customername=T.buyname"   + VbCrlf
	sqlStr = sqlStr + " , opentitle='" + html2db(opentitle) + "'" + VbCrlf
	sqlStr = sqlStr + " , opencontents='" + html2db(opencontents) + "'" + VbCrlf
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master T" + VbCrlf
    sqlStr = sqlStr + " where T.masteridx='" + masteridx + "'"  + VbCrlf
    sqlStr = sqlStr + " and db_shop.dbo.tbl_shopbeasong_cs_master.masteridx=" + CStr(InsertedId)
		
	'response.write sqlStr &"<br>"
	dbget.Execute sqlStr

	''ȸ����û �����ΰ�� - �⺻ ȸ�� ����� ����
	''�±�ȯ, ���� �߼�, �����߼�
	if (divcd="A010") or (divcd="A010") or (divcd="A000") or (divcd="A001") or (divcd="A002") then
	    Call RegDefaultDEliverInfo_off(InsertedId, orderno,masteridx)
    end if

	RegCSMaster_off = InsertedId
end function

''�⺻ ȸ��/�±�ȯ/���񽺹߼� �ּ��� �Է� - ������ �ֹ���ȣ �⺻ �ּ����� �����. - ������ �����ϴ� Procsess
function RegDefaultDEliverInfo_off(AsID, orderno,masteridx)
    dim sqlStr
    
    sqlStr = ""
    sqlStr = "insert into db_shop.dbo.tbl_shopbeasong_cs_delivery"
    sqlStr = sqlStr + " (asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqetcaddr)"    
    sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress"
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master"    
    sqlStr = sqlStr + " where masteridx='" + masteridx + "'"
	
	'response.write sqlStr &"<br>"
    dbget.Execute sqlStr
end function

function GetDefaultTitle_off(divcd, InsertedId, orderno ,masteridx)
    dim opentitle, opencontents ,sqlStr
    dim ipkumdiv, cancelyn, comm_name, ipkumdivName    
	
	sqlStr = ""
	sqlStr = " select m.ipkumdiv,m.cancelyn, C.comm_name"
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
	sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopbeasong_cs_master A"
	sqlStr = sqlStr + "     on A.orderno='" + orderno + "'"
	
	if (masteridx<>"") then
		sqlStr = sqlStr + " and A.masteridx=" + CStr(masteridx)
	end if
	
    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_cs_comm_code_off C"
    sqlStr = sqlStr + " on C.comm_cd='" + divcd + "'"

    sqlStr = sqlStr + " where m.masteridx='" + masteridx + "'"

	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        ipkumdiv    = rsget("ipkumdiv")
        cancelyn    = rsget("cancelyn")
        comm_name   = rsget("comm_name")        
    end if
    
    rsget.close

	GetDefaultTitle_off = comm_name    
end function

function FinishCSMaster_off(iAsid, finishuser, contents_finish)
    dim sqlStr ,IsCsErrStockUpdateRequire

    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate"
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(iAsid)
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        IsCsErrStockUpdateRequire = (rsget("divcd")="A011") and (IsNULL(rsget("finishdate"))) and (rsget("currstate")<>"B007")
    end if
    
    rsget.close

    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master set"	+ VbCrlf
    sqlStr = sqlStr + " finishuser='" + finishuser + "'"            + VbCrlf
    sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(iAsid)

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

    ''�±�ȯȸ�� �Ϸ��ϰ�� ��������Ʈ. 2007.11.16
    'if (IsCsErrStockUpdateRequire) then
    '    sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & iAsid & ",'','" & finishuser & "'"
    '    dbget.Execute sqlStr
    'end if
end function

function AddCustomerOpenContents_off(masteridx, addcontents)
    dim sqlStr

    if ((addcontents="") or (masteridx="")) then Exit Function

    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master set"        + VbCrlf
    sqlStr = sqlStr + " opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & addcontents & "' else '" & VbCrlf & addcontents + "' End )" + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr
end function

dim IsStatusRegister			'����
dim IsStatusEdit				'����
dim IsStatusFinishing			'ó���Ϸ� �õ�
dim IsStatusFinished			'ó���Ϸ�
dim IsDisplayPreviousCSList		'���� CS ����
dim IsDisplayCSMaster			'CS ����������
dim IsDisplayItemList			'��ǰ���
dim IsDisplayRefundInfo			'ȯ������
dim IsDisplayButton				'��ư
dim IsPossibleModifyCSMaster
dim IsPossibleModifyItemList
dim IsPossibleModifyRefundInfo
dim ARR_ERROR_MSG()
dim MAX_ERROR_MSG_COUNT
dim ERROR_MSG_TRY_MODIFY

MAX_ERROR_MSG_COUNT = 10
ReDim Preserve ARR_ERROR_MSG(MAX_ERROR_MSG_COUNT)

'���� ����
function SetCSVariable_off(mode, divcd)
	IsStatusRegister 			= false
	IsStatusEdit 				= false
	IsStatusFinishing 			= false
	IsStatusFinished 			= false
	IsDisplayPreviousCSList 	= true
	IsDisplayCSMaster 			= true
	IsDisplayItemList 			= true
	IsDisplayRefundInfo 		= true
	IsDisplayButton 			= true
	IsPossibleModifyCSMaster	= true
	IsPossibleModifyItemList	= true
	IsPossibleModifyRefundInfo	= true
	
	'CS ����
    if (mode = "regcsas") then	
    	IsStatusRegister 	= true

	'CS ����
    elseif (mode = "editreginfo") then
    	IsStatusEdit 		= true
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS �������¿����� ��ǰ����/ȯ�������� ������ �� �����ϴ�. ������ ���ۼ��ϼ���."
    
    '�Ϸ�õ�
    elseif (mode = "finishreginfo") then
    	IsStatusFinishing 	= true
		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS �Ϸ�ó�� �ܰ迡���� ó�������Է� �� ������ �� �����ϴ�. CS ���������� �̿��ϼ���."
    
    '�Ϸ�� ����
    elseif (mode = "finished") then    	    	
    	IsStatusFinished 	= true
		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false
    	IsDisplayButton 	= false
    	
    	ERROR_MSG_TRY_MODIFY = "�Ϸ�� ������ ������ �� �����ϴ�."
    end if
end function

function GetCSCommName_off(groupCode, divcd)
	dim tmp_str,sqlStr

	sqlStr = " select top 1 comm_cd,comm_name "
	sqlStr = sqlStr + " from  "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_cs_comm_code_off "
	sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
	sqlStr = sqlStr + " and comm_cd='" + CStr(divcd) + "' "
	sqlStr = sqlStr + " and comm_isDel='N' "
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	tmp_str = ""
	if not rsget.EOF  then
		tmp_str = db2html(rsget("comm_name"))
	end if
	rsget.close

	GetCSCommName_off = tmp_str
End function

'�ֹ����
public function IsCSCancelProcess_off(divcd)
	if (divcd = "A008") then
		IsCSCancelProcess_off = true
	else
		IsCSCancelProcess_off = false
	end if
end function

''������ ��ǰ�� üũ ���ɿ���
public function IsPossibleCheckItem_off(divcd, ismastercanceled, isdetailcanceled, masterstate, itemdetailstate, isupchebeasong)
	IsPossibleCheckItem_off = false
	if (ismastercanceled) then exit function
	if (isdetailcanceled) then exit function

	if (IsCSCancelProcess_off(divcd)) then
		IsPossibleCheckItem_off = true
		'/��ǰ�غ��߻���
		if (CStr(itemdetailstate) > "3") then
			IsPossibleCheckItem_off = false
		end if

	elseif (IsCSReturnProcess_off(divcd) = true) then
		IsPossibleCheckItem_off = false
		if (CStr(itemdetailstate) >= "7") then
			if _
				((divcd = "A004") and (isupchebeasong)) _
				or _
				(((divcd = "A010") or (divcd = "A010")) and (Not isupchebeasong)) _
				or _
				(divcd = "A000") _
			then
				'��ǰ����(��ü���)
				'ȸ����û(�ٹ����ٹ��), �±�ȯȸ��(�ٹ����ٹ��)
				'�±�ȯ
				IsPossibleCheckItem_off = true
			end if
		end if
		
	else
		'��Ÿ
		IsPossibleCheckItem_off = true
	end if
end function

'/'��ǰ����(��ü���), ȸ����û(�ٹ����ٹ��), �±�ȯȸ��(�ٹ����ٹ��)
public function IsCSReturnProcess_off(divcd)
	if ((divcd = "A004") or (divcd = "A010") or (divcd = "A011") or (divcd = "A000")) then
		IsCSReturnProcess_off = true
	else
		IsCSReturnProcess_off = false
	end if
end function

function ValidDeleteCS_off(masteridx)
    dim sqlStr
    dim currstate

    ValidDeleteCS_off = false

    sqlStr = "select * from db_shop.dbo.tbl_shopbeasong_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	
	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr,dbget,1
        currstate = rsget("currstate")
    rsget.Close

    If (currstate>="B006") then Exit function

    ValidDeleteCS_off = true
end function

function DeleteCSProcess_off(masteridx, finishuserid)
    dim sqlStr, resultCount

    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master set" + VbCrlf
    sqlStr = sqlStr + "  deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " ,finishuser = '" + finishuserid+ "'" + VbCrlf
    sqlStr = sqlStr + " ,finishdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
    sqlStr = sqlStr + " and currstate<'B006'"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr, resultCount

    DeleteCSProcess_off = (resultCount>0)
end function

'' CS Master ����
function EditCSMaster_off(divcd, orderserial, modiuserid, title, contents_jupsu, csmasteridx)    
    dim sqlStr

    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master"
    sqlStr = sqlStr + " set writeuser='" + modiuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

function EditCSDetailByArrStr_off(byval detailitemlist, csmasteridx, orderno)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dregitemno, dcausecontent

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dregitemno      = tmp(1)
			dcausecontent   = tmp(2)

	        call EditOneCSDetail(csmasteridx, dorderdetailidx, orderno, dregitemno, dcausecontent)
		end if
	next
end function

function EditOneCSDetail(csmasteridx, dorderdetailidx, orderno, dregitemno, dcausecontent)
    dim sqlStr

    sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_detail set"
    sqlStr = sqlStr + " regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and orderdetailidx=" + CStr(dorderdetailidx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

%>