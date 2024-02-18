<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : ���ݰ�꼭 ��������
' History : ������ ����
'			2022.10.31 �ѿ�� ����(���ϰ� ���ݰ�꼭 ���� api �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/cscenter/lib/TaxSheetFunc.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
dim orderserial, issuetype, etcmeachulIdx, taxidx, previssuecount, mode, supplyBusiIdx, busiIdx, IsIssusOK
dim ordercancelyn, chulgoyear, sqlStr, errMSG, i, tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID
dim userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn
dim groupID, itemname, chulgoPrice, taxtype
	IsIssusOK = True
	orderserial 	= requestcheckvar(trim(request("orderserial")),11)
	issuetype 		= trim(request("issuetype"))
	chulgoyear 		= trim(request("chulgoyear"))
	taxidx			= requestcheckvar(getNumeric(trim(request("taxidx"))),10)
	etcmeachulIdx	= trim(request("idx"))		'// ��Ÿ�����ڵ�

previssuecount = 0

dim oTax, oTaxCheck
set oTax = new CTax
oTax.FRecttaxIdx = taxidx

if (taxidx = "") then
	mode = "new"
	oTax.GetTaxEmptyOne

	if (orderserial <> "") or (issuetype = "orderserial") then
		oTax.FOneItem.ForderIdx 	= orderserial
		oTax.FOneItem.Ftaxtype		= "Y"
		oTax.FOneItem.FtotalTax		= 0

		groupID 	= request("groupID")		'// ������ �׷��ڵ�
		itemname 	= request("itemname")
		chulgoPrice = request("chulgoPrice")
		taxtype 	= request("taxtype")
		busiIdx		= request("busiIdx")

		if (chulgoyear <> "") and (chulgoyear >= "2014") and (groupID = "") then
			'// �����û ���
			oTax.FOneItem.Fbilldiv 		= "11"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// �¶���(����)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"
		elseif (groupID <> "") then
			'// ��ü�� ��꼭 ����
			oTax.FOneItem.Fbilldiv 		= "11"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// �¶���(����)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"

			if (groupID <> "G00456") then
				oTax.FOneItem.FconsignYN	= "Y"
				oTax.FOneItem.FissueMethod	= "eSero"
			end if
		else
			'// 2013�⵵ ���� ��꼭 ����
			oTax.FOneItem.Fbilldiv 		= "01"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// �¶���(����)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"
		end if

		oTax.FOneItem.Fuserid		= session("ssBctId")

		'// --------------------------------------------------------------------
		dim ojumun
		set ojumun = new COrderMaster

		if (orderserial <> "") then
			ojumun.FRectOrderSerial = orderserial
			ojumun.QuickSearchOrderMaster
		end if

		if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
			ojumun.FRectOldOrder = "on"
			ojumun.QuickSearchOrderMaster
		end if

		if (ojumun.FResultCount < 1) and (errMSG = "") then
			errMSG = "�߸��� �ֹ���ȣ �Դϴ�."
		else
			oTax.FOneItem.FisueDate = getMayTaxDate(ojumun.FOneItem.Fipkumdate)
		end if

		'// --------------------------------------------------------------------
		dim oTaxPrevIssue
		set oTaxPrevIssue = new CTax

		if (errMSG = "") then
			oTax.FCurrPage = 1
			oTax.FPageSize = 100
			''oTax.FRectsearchBilldiv = "01"				'�Һ��ڸ���
			oTax.FRectsearchKey = "t.userid"
			oTax.FRectDelYn = "N"

			if (ojumun.FOneItem.FUserID <> "") then
				oTax.FRectsearchString = ojumun.FOneItem.FUserID
				userid = ojumun.FOneItem.FUserID
			else
				oTax.FRectsearchString = "----"
			end if

			oTax.GetTaxList

			previssuecount = oTax.FTotalCount
		end if

		''����� ���ݰ�꼭���� üũ
		if errMSG = "" and (oTax.FOneItem.ForderIdx <> "") then
			set oTaxCheck = new CTax

			oTaxCheck.FRectsearchKey = " t.orderserial "
			oTaxCheck.FRectsearchString = CStr(oTax.FOneItem.ForderIdx)
			oTaxCheck.FRectDelYn = "N"

			if (groupID <> "") then
				oTaxCheck.FRectSupplyGroupID = groupID
			end if

			oTaxCheck.GetTaxList

			if oTaxCheck.FResultCount > 0 and (oTax.FOneItem.Fbilldiv = "01" or groupID <> "") then
				if oTaxCheck.FTaxList(0).FisueYn="Y" then
					if (errMSG = "") then
						errMSG = "�̹� ����� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� �ŷ�ó�� [��ҿ�û]�� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�"
					end if
				else
					if (errMSG = "") then
						errMSG = "���������� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�"
					end if
				end if
			end if
		end if

		sqlStr =	"select ( select " &_
				"			Case " &_
				"				When count(idx)>1 Then max(itemname) + '�� ' + Cast((count(idx)-1) as varchar) + '��' " &_
				"				Else max(itemname) " &_
				"			End " &_
				"		from db_order.[dbo].tbl_order_detail " &_
				"		where orderserial='" & orderserial & "' and itemid<>0 and cancelyn='N' group by orderserial " &_
				"	) as itemname " &_
				"	, subtotalprice, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " &_
				"from db_order.[dbo].tbl_order_master " &_
				"Where orderserial = '" & orderserial & "'"

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if Not(rsget.EOF or rsget.BOF) then
			oTax.FOneItem.Fitemname = rsget("itemname")

			if (chulgoyear <> "") and (chulgoyear >= "2014") and (groupID <> "") then
				oTax.FOneItem.Fitemname = "��ü�� ��ǰ"

				if request("itemname") <> "" then
					oTax.FOneItem.Fitemname = request("itemname")
					oTax.FOneItem.FtotalPrice = request("chulgoPrice")
					oTax.FOneItem.Ftaxtype = request("taxtype")
				end if
			elseif (CStr(rsget("accountdiv")) = "7") or (CStr(rsget("accountdiv")) = "20") then
				'������, �ǽð���ü : ��ü�ݾ�
				oTax.FOneItem.FtotalPrice = rsget("subtotalprice")
			else
				'���������ݾ׸�
				oTax.FOneItem.FtotalPrice = rsget("sumPaymentEtc")
			end if
		end if
		rsget.close

		if (groupID <> "") then
			dim ogroupSupply
			set ogroupSupply = new CPartnerGroup
			ogroupSupply.FRectGroupid = groupID
			ogroupSupply.GetOneGroupInfo

			if (groupID = "G00456") then
				'// ������(�ٹ�����)
				oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
				oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
				oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
				oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
				oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
				oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
				oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
				oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
				oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL
			else
				'// ������(��ü)
				oTax.FOneItem.FsupplybusiNo			= ogroupSupply.FOneItem.Fcompany_no
				oTax.FOneItem.FsupplybusiName		= ogroupSupply.FOneItem.FCompany_name
				oTax.FOneItem.FsupplybusiCEOName	= ogroupSupply.FOneItem.Fceoname
				oTax.FOneItem.FsupplybusiAddr		= ogroupSupply.FOneItem.Fcompany_address & " " & ogroupSupply.FOneItem.Fcompany_address2
				oTax.FOneItem.FsupplybusiType		= ogroupSupply.FOneItem.Fcompany_uptae
				oTax.FOneItem.FsupplybusiItem		= ogroupSupply.FOneItem.Fcompany_upjong
				oTax.FOneItem.FsupplyrepName		= ogroupSupply.FOneItem.Fjungsan_name
				oTax.FOneItem.FsupplyrepTel			= ogroupSupply.FOneItem.Fjungsan_hp
				oTax.FOneItem.FsupplyrepEmail		= ogroupSupply.FOneItem.Fjungsan_email
			end if


		end if

		if (busiIdx <> "") then
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " 	b.busiNo, b.busiSubNo, b.busiName, b.busiCEOName, b.busiAddr, b.busiType, b.busiItem, b.repName, b.repEmail, b.repTel "
			sqlStr = sqlStr + " from db_order.dbo.tbl_busiInfo b "
			sqlStr = sqlStr + " where b.busiidx = " + CStr(busiIdx) + " "
			rsget.Open sqlStr, dbget, 1

			if Not(rsget.EOF or rsget.BOF) then
				'// ���޹޴���
				oTax.FOneItem.FbusiNo = rsget("busiNo")
				oTax.FOneItem.FbusiSubNo = rsget("busiSubNo")
				oTax.FOneItem.FbusiName = rsget("busiName")
				oTax.FOneItem.FbusiCEOName = rsget("busiCEOName")
				oTax.FOneItem.FbusiAddr = rsget("busiAddr")
				oTax.FOneItem.FbusiType = rsget("busiType")
				oTax.FOneItem.FbusiItem = rsget("busiItem")
				oTax.FOneItem.FrepName = rsget("repName")
				oTax.FOneItem.FrepEmail = rsget("repEmail")
				oTax.FOneItem.FrepTel = rsget("repTel")
			end if
			rsget.close
		end if

	end if

	if (issuetype = "etcmeachul") and (etcmeachulIdx <> "") then

		'// --------------------------------------------------------------------
		'// ��Ÿ����
		dim oetcmeachul
		set oetcmeachul = new CEtcMeachul
		oetcmeachul.FRectidx = etcmeachulIdx
		oetcmeachul.getOneEtcMeachul

		oTax.FOneItem.FsellBizCd	= oetcmeachul.FOneItem.Fbizsection_cd
		oTax.FOneItem.Fselltype		= oetcmeachul.FOneItem.Fselltype
		oTax.FOneItem.Ftaxissuetype	= "E"
		oTax.FOneItem.Ftaxtype		= "Y"
		oTax.FOneItem.Fbilldiv		= "51"
		oTax.FOneItem.FconsignYN	= "N"
		oTax.FOneItem.FissueMethod	= "WEHAGO"

		oTax.FOneItem.Fuserid		= session("ssBctId")

		'// --------------------------------------------------------------------
		'����̵𿡼� �׷��ڵ� ����
		dim opartner
		set opartner = new CPartnerUser

		opartner.FCurrpage = 1
		opartner.FPageSize = 100
		opartner.FRectDesignerID = oetcmeachul.FOneItem.Fshopid
		opartner.FRectUserDiv = "all"

		opartner.GetPartnerNUserCList

		'// --------------------------------------------------------------------
		'�׷��ڵ忡�� ���ݰ�꼭/�������� ���� ����
		dim ogroup
		set ogroup = new CPartnerGroup

		if (opartner.FResultCount > 0) then
			ogroup.FRectGroupid = opartner.FPartnerList(0).FGroupID
			ogroup.GetOneGroupInfo
		else
			ogroup.FResultCount = 0
		end if

		if (opartner.FResultCount < 1) then
			errMSG = "�߸��� �귣���Դϴ�."
		elseif (ogroup.FResultCount < 1) then
			errMSG = "�׷��ڵ尡 �����Ǿ� ���� ���� ��ü�Դϴ�."
		else
			'// ���޹޴���
			if (ogroup.FOneItem.Fcompany_no <> "211-87-00620") then
				if ogroup.FOneItem.fBIZ_NO<>"" then
					oTax.FOneItem.FbusiNo		= ogroup.FOneItem.fBIZ_NO
				else					
					oTax.FOneItem.FbusiNo		= ogroup.FOneItem.Fcompany_no
				end if
				if ogroup.FOneItem.fCUST_NM<>"" then
					oTax.FOneItem.FbusiName		= ogroup.FOneItem.fCUST_NM
				else					
					oTax.FOneItem.FbusiName		= ogroup.FOneItem.FCompany_name
				end if
				if ogroup.FOneItem.fCEO_NM<>"" then
					oTax.FOneItem.FbusiCEOName	= ogroup.FOneItem.fCEO_NM
				else					
					oTax.FOneItem.FbusiCEOName	= ogroup.FOneItem.Fceoname
				end if
				if ogroup.FOneItem.faddr<>"" then
					oTax.FOneItem.FbusiAddr		= ogroup.FOneItem.faddr		' ogroup.FOneItem.fPOST_CD
				else					
					oTax.FOneItem.FbusiAddr		= ogroup.FOneItem.Fcompany_address & " " & ogroup.FOneItem.Fcompany_address2
				end if
				if ogroup.FOneItem.fBSCD<>"" then
					oTax.FOneItem.FbusiType		= ogroup.FOneItem.fBSCD
				else					
					oTax.FOneItem.FbusiType		= ogroup.FOneItem.Fcompany_uptae
				end if
				if ogroup.FOneItem.fINTP<>"" then
					oTax.FOneItem.FbusiItem		= ogroup.FOneItem.fINTP
				else					
					oTax.FOneItem.FbusiItem		= ogroup.FOneItem.Fcompany_upjong
				end if
				
				oTax.FOneItem.FrepName		= ogroup.FOneItem.Fjungsan_name

				if ogroup.FOneItem.fTEL_NO<>"" then
					oTax.FOneItem.FrepTel		= ogroup.FOneItem.fTEL_NO
				else					
					oTax.FOneItem.FrepTel		= ogroup.FOneItem.Fjungsan_hp
				end if
				if ogroup.FOneItem.fEMAIL<>"" then
					oTax.FOneItem.FrepEmail		= ogroup.FOneItem.fEMAIL
				else					
					oTax.FOneItem.FrepEmail		= ogroup.FOneItem.Fjungsan_email
				end if			
			end if

			oTax.FOneItem.FconfirmYn	= "Y"
			oTax.FOneItem.FtotalPrice	= oetcmeachul.FOneItem.Ftotalsum
			oTax.FOneItem.FtotalTax		= 0
			oTax.FOneItem.Fitemname		= oetcmeachul.FOneItem.Ftitle
			oTax.FOneItem.ForderIdx		= etcmeachulIdx
		end if

		'// --------------------------------------------------------------------
		'' ����̵𿡼� 3PL ��ü���� Ȯ��
		if (Is3PLShopid(oetcmeachul.FOneItem.Fshopid) = True) then
			Call Get3PLUpcheInfoByShopid(oetcmeachul.FOneItem.Fshopid, tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID)

			dim otplgroup
			set otplgroup = new CPartnerGroup

			otplgroup.FRectGroupid = tplgroupid
			otplgroup.GetOneGroupInfo

			if (otplgroup.FResultCount < 1) then
				errMSG = "3PL�׷��ڵ尡 �����Ǿ� ���� ���� ��ü�Դϴ�."
			else
				'// ������(3PL)
				oTax.FOneItem.FsupplyBusiNo			= otplgroup.FOneItem.Fcompany_no
				oTax.FOneItem.FsupplyBusiName		= otplgroup.FOneItem.FCompany_name
				oTax.FOneItem.FsupplyBusiCEOName	= otplgroup.FOneItem.Fceoname
				oTax.FOneItem.FsupplyBusiAddr		= otplgroup.FOneItem.Fcompany_address & " " & otplgroup.FOneItem.Fcompany_address2
				oTax.FOneItem.FsupplyBusiType		= otplgroup.FOneItem.Fcompany_uptae
				oTax.FOneItem.FsupplyBusiItem		= otplgroup.FOneItem.Fcompany_upjong
				oTax.FOneItem.FsupplyRepName		= otplgroup.FOneItem.Fjungsan_name
				oTax.FOneItem.FsupplyRepTel			= otplgroup.FOneItem.Fjungsan_hp
				oTax.FOneItem.FsupplyRepEmail		= otplgroup.FOneItem.Fjungsan_email

				oTax.FOneItem.Fbilldiv = "99"
				oTax.FOneItem.FsupplyConfirmYn = "Y"
			end if
		else
			'// ������(�ٹ�����)
			oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
			oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
			oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
			oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
			oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
			oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
			oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
			oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
			oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL

			oTax.FOneItem.FsupplyConfirmYn = "Y"
		end if

		''����� ���ݰ�꼭���� üũ
		if errMSG = "" and (oTax.FOneItem.ForderIdx <> "") then
			set oTaxCheck = new CTax

			oTaxCheck.FRectsearchKey = " t.orderidx "
			oTaxCheck.FRectsearchString = CStr(oTax.FOneItem.ForderIdx)
			oTaxCheck.FRectDelYn = "N"

			oTaxCheck.GetTaxList

			if oTaxCheck.FResultCount > 0 then
				if oTaxCheck.FTaxList(0).FisueYn="Y" then
					if (errMSG = "") then
						errMSG = "�̹� ����� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� �ŷ�ó�� [��ҿ�û]�� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�"
					end if
				else
					if (errMSG = "") then
						errMSG = "���������� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�"
					end if
				end if
			end if
		end if
	end if

	if (oTax.FOneItem.FTaxType = "") then
		oTax.FOneItem.FTaxType = "Y"
	end if

	if (oTax.FOneItem.FsupplyBusiNo = "") then
		''oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
		''oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
		''oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
		''oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
		''oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
		''oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
		''oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
		''oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
		''oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL

		oTax.FOneItem.FsupplyConfirmYn = "Y"
	end if
else
	mode = "view"
	oTax.GetTaxRead

	'response.write oTax.FOneItem.Fbilldiv
	'response.write oTax.FOneItem.Ftplcompanyid

	if (oTax.FOneItem.Fbilldiv = "99") then
		if Not IsNull(oTax.FOneItem.Ftplcompanyid) then
			Call Get3PLUpcheInfo(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID)
		end if
	end if
end if

if (errMSG <> "") then
	IsIssusOK = False
end if

'���ͺμ����
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FSale = "N"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>
<script type='text/javascript'>

var errMSG = "<%= errMSG %>";

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function setRegisterInfo() {
	var TENBYTEN_SOCNO = "<%= TENBYTEN_SOCNO %>";
	var TENBYTEN_SOCNAME = "<%= TENBYTEN_SOCNAME %>";
	var TENBYTEN_CEONAME = "<%= TENBYTEN_CEONAME %>";
	var TENBYTEN_SOCADDR = "<%= TENBYTEN_SOCADDR %>";
	var TENBYTEN_SOCSTATUS = "<%= TENBYTEN_SOCSTATUS %>";
	var TENBYTEN_SOCEVENT = "<%= TENBYTEN_SOCEVENT %>";
	var TENBYTEN_MANAGERNAME = "<%= TENBYTEN_MANAGERNAME %>";
	var TENBYTEN_MANAGERPHONE = "<%= TENBYTEN_MANAGERPHONE %>";
	var TENBYTEN_MANAGERMAIL = "<%= TENBYTEN_MANAGERMAIL %>";

	var SUPPLY_SOCNO = "<%= oTax.FOneItem.FsupplybusiNo %>";
	var SUPPLY_SOCNAME = "<%= oTax.FOneItem.FsupplybusiName %>";
	var SUPPLY_CEONAME = "<%= oTax.FOneItem.FsupplybusiCEOName %>";
	var SUPPLY_SOCADDR = "<%= oTax.FOneItem.FsupplybusiAddr %>";
	var SUPPLY_SOCSTATUS = "<%= oTax.FOneItem.FsupplybusiType %>";
	var SUPPLY_SOCEVENT = "<%= oTax.FOneItem.FsupplybusiItem %>";
	var SUPPLY_MANAGERNAME = "<%= oTax.FOneItem.FsupplyrepName %>";
	var SUPPLY_MANAGERPHONE = "<%= oTax.FOneItem.FsupplyrepTel %>";
	var SUPPLY_MANAGERMAIL = "<%= oTax.FOneItem.FsupplyrepEmail %>";

	// 01, 02, 03, 51, 11, 99
	if ((frm.billdiv.value == "01") || (frm.billdiv.value == "02") || (frm.billdiv.value == "03") || (frm.billdiv.value == "51")) {
		setSupplyCompanyInfo(TENBYTEN_SOCNO, "", TENBYTEN_SOCNAME, TENBYTEN_CEONAME, TENBYTEN_SOCADDR, TENBYTEN_SOCSTATUS, TENBYTEN_SOCEVENT, TENBYTEN_MANAGERNAME, TENBYTEN_MANAGERPHONE, TENBYTEN_MANAGERMAIL);
	} else if (frm.billdiv.value == "11") {
		setSupplyCompanyInfo(SUPPLY_SOCNO, "", SUPPLY_SOCNAME, SUPPLY_CEONAME, SUPPLY_SOCADDR, SUPPLY_SOCSTATUS, SUPPLY_SOCEVENT, SUPPLY_MANAGERNAME, SUPPLY_MANAGERPHONE, SUPPLY_MANAGERMAIL);
	}
}

function chgHandTax(comp){
    var txbox = comp.form.totaltaxprice;

    if (comp.checked){
        txbox.readOnly = false;
        txbox.className = "writebox";
    }else{
        txbox.readOnly = true;
        txbox.className = "readonlybox";
    }
}

function CalcPrice(comp) {
	var frm = document.frm;

	if (frm.taxtype.value.length < 1) {
		alert('���������� �Է��ϼ���.');
		return;
	}

	if ((comp.name == "taxtype") || (comp.name == "totalprice")) {
		if ((frm.totalprice.value == "") || (frm.totalprice.value*0 != 0)) {
			// alert("���� �հ�ݾ��� �Է��ϼ���.");
			return;
		}

		if (frm.taxtype.value == "Y") {
			// ������ ���ް��� ���ϰ� 0.1 �� �ݿø� ���ָ� �ȴ�.
			frm.totaltaxprice.value = Math.round(1.0 * frm.totalprice.value / 1.1 / 10.0);
			frm.totaltaxprice2.value = frm.totaltaxprice.value;
		} else {
			frm.totaltaxprice.value = 0;
		}
		frm.totaltaxprice2.value = frm.totaltaxprice.value;

		frm.totalsupplyprice.value = frm.totalprice.value*1 - frm.totaltaxprice2.value*1;
		frm.totalsupplyprice2.value = frm.totalsupplyprice.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice.value;
	} else if (comp.name == "totaltaxprice") {
		if ((frm.totaltaxprice.value == "") || (frm.totaltaxprice.value*0 != 0)) {
			// alert("���� ������ �Է��Է��ϼ���.");
			return;
		}

		frm.totaltaxprice2.value = frm.totaltaxprice.value;

		frm.totalsupplyprice.value = frm.totalprice.value*1 - frm.totaltaxprice2.value*1;
		frm.totalsupplyprice2.value = frm.totalsupplyprice.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice.value;
	} else if (comp.name == "totalsupplyprice2") {
		if ((frm.totalsupplyprice2.value == "") || (frm.totalsupplyprice2.value*0 != 0)) {
			// alert("���� ������ �Է��Է��ϼ���.");
			return;
		}

		frm.totalsupplyprice.value = frm.totalsupplyprice2.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice2.value;

		if (frm.taxtype.value == "Y") {
			// ������ ���ް��� ���ϰ� 0.1 �� �ݿø� ���ָ� �ȴ�.
			frm.totaltaxprice.value = parseInt(frm.totalsupplyprice2.value*0.1);
			frm.totaltaxprice2.value = frm.totaltaxprice.value;
		} else {
			frm.totaltaxprice.value = 0;
		}

		frm.totalprice.value = frm.totalsupplyprice.value*1 + frm.totaltaxprice.value*1;
	}

	frm.totalprice2.value = frm.totalprice.value;
	frm.totalprice3.value = frm.totalprice.value;
}

function doRegisterSheet(){

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if (frm.issuetype.value != "") {
		if ((frm.issuetype.value == "orderserial") && (frm.billdiv.value != "01") && (frm.billdiv.value != "11")) {
			alert('�Һ��� ���⸸ �ۼ��� �� �ֽ��ϴ�.');
			frm.billdiv.focus();
			return;
		}

		if ((frm.issuetype.value == "etcmeachul") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
			alert('��Ÿ���⸸ �ۼ��� �� �ֽ��ϴ�.');
			frm.billdiv.focus();
			return;
		}

		if(frm.orderidx.value == "") {
			alert('�ֹ���ȣ �Ǵ� ��Ÿ���� �ڵ尡 ��� �ԷµǾ� �־�� �մϴ�.');
			return;
		}
	}

	// 20036 => 4010005
	if ((frm.selltype.value == "4010005") && (frm.taxtype.value != "0")) {
		alert('���������� ������ ��� ������꼭�� �ۼ������մϴ�.');
		return;
	} else if ((frm.selltype.value != "4010005") && (frm.taxtype.value == "0")) {
		alert('���������� ������ �ƴϸ� ������꼭�� �ۼ��� �� �����ϴ�.');
		return;
	}

	if(frm.billdiv.value == "0") {
		alert('�����ڸ� �����ϼ���.');
		return;
	}

	if (frm.socname.value.length<1){
		alert('����� ��ϻ��� ȸ����� �Է��ϼ���.');
		frm.socname.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('����� ��ϻ��� ��ǥ�ڸ��� �Է��ϼ���.');
		frm.ceoname.focus();
		return;
	}

	if (frm.socno.value.length<1){
		alert('����� ��� ��ȣ�� �Է��ϼ���.');
		frm.socno.focus();
		return;
	}

	if (frm.socaddr.value.length<1){
		alert('����� ��ϻ��� �ּҸ� �Է��ϼ���.');
		frm.socaddr.focus();
		return;
	}

	if (frm.socstatus.value.length<1){
		alert('����� ��ϻ��� ���¸� �Է��ϼ���.');
		frm.socstatus.focus();
		return;
	}

	if (frm.socevent.value.length<1){
		alert('����� ��ϻ��� ������ �Է��ϼ���.');
		frm.socevent.focus();
		return;
	}

	if (frm.managername.value.length<1){
		alert('����� ������ �Է��ϼ���.');
		frm.managername.focus();
		return;
	}

	if (frm.managerphone.value.length<1){
		alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managerphone.value.indexOf("-") == -1) {
		alert('����� ����ó�� 000-000-0000 �����̾�� �մϴ�.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managermail.value.length<1){
		alert('����� �̸����ּҸ� �Է��ϼ���.');
		frm.managermail.focus();
		return;
	}

	if (frm.yyyymmdd.value.length<1){
		alert('�ۼ����� �Է��ϼ���.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('ǰ���� �Է��ϼ���.');
		return;
	}

	if (frm.totalprice.value.length<1){
		alert('�հ�ݾ��� �Է��ϼ���.');
		return;
	}

	if (frm.taxtype.value.length<1){
		alert('���������� �Է��ϼ���.');
		return;
	}

	if ((frm.subsocno.value.length != 0) && (frm.subsocno.value.length != 4)) {
		alert('��������ȣ�� 4�ڸ��� �Է��ϼ���');
		return;
	}

	if ((frm.billdiv.value == "01") || (frm.billdiv.value == "11")) {
		if(frm.orderidx.value == "") {
			alert('��� �ֹ���ȣ�� �Է��ϼ���.');
			return;
		}
	} else if ((frm.orderidx.value != "") && (frm.billdiv.value != "03") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
		alert('�Һ��ڸ���/���θ��/��Ÿ���⿡�� ��� �ֹ���ȣ �Ǵ� ����ڵ带 ���� �� �ֽ��ϴ�.');
		return;
	}

	if (frm.billdiv.value != "99") {
		if (frm.sellBizCd.value.length<1){
			alert('����μ��� �����ϼ���.');
			return;
		}
	} else {
		if (frm.sellBizCd.value.length > 0){
			alert('3PL���⿡�� �μ��� ������ �� �����ϴ�.');
			return;
		}
	}

	if (frm.selltype.value.length<1){
		alert('��������� �����ϼ���.');
		return;
	}

	if (frm.taxissuetype.value.length<1){
		alert('���γ����� �����ϼ���.');
		return;
	}

	<% if (groupID <> "") then %>
	if (frm.billdiv.value == "11") {
		// �Һ���(��ü����)
		if ((frm.consignYN.value == "N") && (frm.reg_socno.value != "211-87-00620")) {
			alert("�����ڰ� �ٹ������� �ƴ� ��� ����Ź���� �����ϼ���.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.reg_socno.value == "211-87-00620")) {
			alert("�����ڰ� �ٹ������� ��� ����Ź������ [����] ���� �����ϼ���.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.issueMethod.value != "eSero")) {
			alert("����Ź ��꼭�� �̼��ο��� ������ุ ���డ���մϴ�.");
			return;
		}
	}
	<% end if %>

	setRegisterInfo();

    if (confirm('���ݰ�꼭 �����û�� �Ͻðڽ��ϱ�?')){
		document.frm.mode.value = "tax_register_new";

		if (frm.billdiv.value == "11") {
			<% if (groupID <> "") then %>
			// 2014�� ���� ��꼭 ����(��ü��)
			document.frm.mode.value = "tax_register_new_2014_upche";
			<% else %>
			// 2014�� ���� ��û(��ü��)
			document.frm.mode.value = "tax_register_new_2014";
			<% end if %>
		}

        document.frm.submit();
    }
}

function popMeachulDetailList() {
	if (frm.taxissuetype.value != "E") {
		alert("��Ÿ������ ��쿡�� ������ �߰��� �� �ֽ��ϴ�.");
		return;
	}

	var popwin = window.open('pop_etc_meachul_list.asp?idx=<%= oTax.FOneItem.ForderIdx %>','popMeachulDetailList','width=1000, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popListPreviousCustomerTaxSheet(userid){
    var popwin=window.open("/cscenter/taxsheet/popListPreviousCustomerTaxSheet.asp?userid=" + userid,"popListPreviousCustomerTaxSheet","width=700,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ReactMeachulDetailList(arrchk, tottaxsum) {
    var frm = document.frm;

    frm.totalprice.value = tottaxsum;
    frm.orderidx.value = arrchk;

    CalcPrice(frm.totalprice);
}

function SearchSocno() {
	if (frm.socno.value == "") {
		alert("����ڹ�ȣ�� �Է��ϼ���.");
		return;
	}

	if (frm.socno.value.length != 12) {
		alert("����ڹ�ȣ�� �Ʒ��� ���� �������� �Է��ϼ���.\n\n000-00-00000");
		return;
	}

	icheckframe.location.href="isearchframe.asp?socno=" + frm.socno.value;
}

function setCompanyInfo(subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
{
	frm.subsocno.value = subsocno;
	frm.socname.value = socname;
	frm.ceoname.value = ceoname;
	frm.socaddr.value = socaddr;
	frm.socstatus.value = socstatus;
	frm.socevent.value = socevent;
	frm.managername.value = managername;
	frm.managerphone.value = managerphone;
	frm.managermail.value = managermail;
}

function setSupplyCompanyInfo(socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
{
	frm.reg_socno.value = socno;
	frm.reg_subsocno.value = subsocno;
	frm.reg_socname.value = socname;
	frm.reg_ceoname.value = ceoname;
	frm.reg_socaddr.value = socaddr;
	frm.reg_socstatus.value = socstatus;
	frm.reg_socevent.value = socevent;
	frm.reg_managername.value = managername;
	frm.reg_managerphone.value = managerphone;
	frm.reg_managermail.value = managermail;
}

// ��û�� ����
function GotoTaxModify(){
	if (frm.managername.value.length<1){
		alert('����� ������ �Է��ϼ���.');
		frm.managername.focus();
		return;
	}

	if (frm.managerphone.value.length<1){
		alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managermail.value.length<1){
		alert('����� �̸����ּҸ� �Է��ϼ���.');
		frm.managermail.focus();
		return;
	}

	if (frm.itemname.value.length<1){
		alert('ǰ���� �Է��ϼ���.');
		return;
	}

	if ((frm.billdiv.value == "52") || (frm.billdiv.value == "55")) {
		alert('����Ұ�');
		return;
	}

	if (frm.totalsupplyprice2.value.length<1){
		alert('�ܰ��� �Է��ϼ���.');
		return;
	}

	if (frm.totalprice.value.length<1){
		alert('�հ�ݾ��� �Է��ϼ���.');
		return;
	}

	if (frm.totalprice.value*1 != frm.orgtotalprice.value*1) {
		alert('�ݾ��� ������ �� �����ϴ�.\n\n���� �� ���ۼ��ϼ���.');
		return;
	}

	if (frm.sellBizCd.value.length<1){
		alert('����μ��� �����ϼ���.');
		return;
	}

	if (frm.selltype.value.length<1){
		alert('��������� �����ϼ���.');
		return;
	}

	if (frm.taxissuetype.value.length<1){
		alert('���γ����� �����ϼ���.');
		return;
	}

	if (frm.taxtype.value.length<1){
		alert('���������� �Է��ϼ���.');
		return;
	}

	if (frm.managerphone.value.indexOf("-") == -1) {
		alert('����� ����ó�� 000-000-0000 �����̾�� �մϴ�.');
		frm.managerphone.focus();
		return;
	}

	if (frm.billdiv.value == "11") {
		// �Һ���(��ü����)
		if ((frm.consignYN.value == "N") && (frm.reg_socno.value != "211-87-00620")) {
			alert("�����ڰ� �ٹ������� �ƴ� ��� ����Ź���� �����ϼ���.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.reg_socno.value == "211-87-00620")) {
			alert("�����ڰ� �ٹ������� ��� ����Ź������ [����] ���� �����ϼ���.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.issueMethod.value != "eSero")) {
			alert("����Ź ��꼭�� �̼��ο��� ������ุ ���డ���մϴ�.");
			return;
		}
	}

	setRegisterInfo();

	if (confirm('��û���� ���� �Ͻðڽ��ϱ�?')){
		document.frm.mode.value="tax_modify";
		document.frm.submit();
	}
}

//function PopCommonSampleTaxReg(){
//	var popCommonSamplewin = window.open("/cscenter/taxsheet/popCommonSampleWehagotaxregapi.asp?taxIdx=<%'=taxIdx %>&taxType=<%'= oTax.FOneItem.Fbilldiv %>","popCommonSampletaxreg","width=1200 height=768 scrollbars=yes resizable=yes");
//	popCommonSamplewin.focus();
//}

function PopCommonWehagoTaxReg(){
	<% if (IsIssusOK = False) then %>
		<% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
			if (!confirm('<%= ErrMSG %>\n\n����Ͻðڽ��ϱ�?(�����ڱ���)')) return;
		<% else %>
			alert('<%= ErrMSG %>');
			return;
		<% end if %>
	<% end if %>

	<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// �������̻� ���� �Ǵ� �濵������, �����(���̶��)
		if (confirm('���ݰ�꼭�� �����Ͻðڽ��ϱ�?')){
			var popCommonWehagowin = window.open('/cscenter/taxsheet/popCommonWehagotaxregapi.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','popCommonWehagotaxreg','width=1200,height=768,scrollbars=yes,resizable=yes');
			popCommonWehagowin.focus()
		}
	<% else %>
		alert('������ �����ϴ�.[2]');
	<% end if %>
}

/*
function TaxEvalBill36524api(){
	<% if (IsIssusOK = False) then %>
		<% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
			if (!confirm('<%= ErrMSG %>\n\n����Ͻðڽ��ϱ�?(�����ڱ���)')) return;
		<% else %>
			alert('<%= ErrMSG %>');
			return;
		<% end if %>
	<% end if %>

	<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// �������̻� ���� �Ǵ� �濵������, �����(���̶��)
		if (confirm('���ݰ�꼭�� �����Ͻðڽ��ϱ�?')){
			var popwin = window.open('/cscenter/taxsheet/evalTaxBill36524api.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','evalTaxBill36524api','width=1024,height=768,scrollbars=yes,resizable=yes');
			popwin.focus()
		}
	<% else %>
		alert('������ �����ϴ�.[2]');
	<% end if %>
}

    function TaxEvalBill36524(){
    	<% if (IsIssusOK = False) then %>
    	    <% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
    	    if (!confirm('<%= ErrMSG %>\n\n����Ͻðڽ��ϱ�?(�����ڱ���)')) return;
    	    <% else %>
    		alert('<%= ErrMSG %>');
    		return;
    		<% end if %>
    	<% end if %>

<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// �������̻� ���� �Ǵ� �濵������, �����(���̶��)
        if (confirm('���ݰ�꼭�� �����Ͻðڽ��ϱ�?')){
            var popwin = window.open('evalTaxBill36524.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','evalTaxBill36524','width=400,height=300,scrollbars=yes,resizable=yes');
            popwin.focus()
        }
<% else %>
		alert('������ �����ϴ�.[2]');
<% end if %>
    }
*/

    function TaxEvaleSero() {
    	<% if (IsIssusOK = False) then %>
    	    <% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
    	    if (!confirm('<%= ErrMSG %>\n\n����Ͻðڽ��ϱ�?(�����ڱ���)')) return;
    	    <% else %>
    		alert('<%= ErrMSG %>');
    		return;
    		<% end if %>
    	<% end if %>

<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// �������̻� ���� �Ǵ� �濵������, �����(���̶��)
        if (confirm('������� �Ϸ�ó�� �Ͻðڽ��ϱ�?')){
			document.frm.mode.value="finishSheetByESero";
			document.frm.submit();
        }
<% else %>
		alert('������ �����ϴ�.[2]');
<% end if %>
    }

	// ��û�� ����
	function GotoTaxDel(isIssued){
		if (confirm('��û���� ���� �Ͻðڽ��ϱ�?\n\n��꼭�� ����� ��� ������ ��ҵ��� �����Ͻñ� �ٶ��ϴ�.')){
		    if (isIssued == "Y") {
    		    <% if C_ADMIN_AUTH or C_MngPowerUser then %>
    		    alert('������ ���� ����. �ݵ�� �������ݰ�꼭 Ȯ��.');
    			document.frm.mode.value="sheetDel";
    			document.frm.submit();
    			<% else %>
    			alert('������ �����ϴ�. ������ ���� ���[1]');
    			<% end if %>
    		}else{
    		    document.frm.mode.value="sheetDel";
    			document.frm.submit();
    		}
		}
	}

	// ���ݰ�꼭 ����
	function goView(tax_no, b_biz_no, s_biz_no)
	{
		<% if (application("Svr_Info")="Dev") then %>
			// �׽�Ʈ
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// �Ǽ���
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no="+b_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView2(tax_no, b_biz_no, s_biz_no){
		<% if (application("Svr_Info")="Dev") then %>
			// �׽�Ʈ
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// �Ǽ���
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView_Bill36524(tax_no, b_biz_no)
	{
			window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	}

    function GotoTaxMapHand(){
        var popwin = window.open('popTaxMapHand.asp?taxIdx=<%=taxIdx%>','popTaxMapHand','scrollbars=yes,resizable=yes,width=400,height=300');
        popwin.focus();
    }

	// ����ڰ˻�
	function PopUpcheSelectBySocno(frmname){
		var socno = eval(frmname+'.socno').value;
		if (socno==''){
			alert('����� ��Ϲ�ȣ�� �Է��ϼ���');
			eval(frmname+'.socno').focus();
			return;
		}

		var popwin = window.open("/admin/member/popupcheselect.asp?mode=tax&frmname=" + frmname + "&rectsocno=" + socno,"popupcheselectbysocno","width=1280 height=960 scrollbars=yes resizable=yes");
		popwin.focus();
	}

	function jsGetCust(frmname){
		var socno = eval(frmname+'.socno').value;
		if (socno==''){
			alert('����� ��Ϲ�ȣ�� �Է��ϼ���');
			eval(frmname+'.socno').focus();
			return;
		}

		var Strparm = "";
		Strparm = "?selSTp=5&sSTx="+ socno;
		Strparm = Strparm + "&opnType=eTaxdetail";
		var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1280, height=960,resizable=yes, scrollbars=yes");
		winC.focus();
	}

	//�ŷ�ó ����
	function jsSetCust(custcd, custnm, ceonm, custno, addr, bscd, intp, email, telno){
		frm.socname.value = custnm;		// ��ȣ
		frm.ceoname.value = ceonm;		// ��ǥ��
		//frm.socno.value = custno;		// ����ڹ�ȣ
		frm.socaddr.value = addr;		// ������ּ�
		frm.socstatus.value = bscd;		// ����
		frm.socevent.value = intp;		// ����
		frm.managerphone.value = telno;		// ����ó
		frm.managermail.value = email;		// �̸���
	}

</script>

<style type="text/css">
.Readonlybox { border:0px; }
.writebox { border:10px; background:#E6E6E6; }
</style>

<table width="800" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<!-- ���ݰ�� ��û�� ���� ���� -->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td colspan="4" align="left">
					<b>���ݰ�꼭 ��������</b>
				</td>
			</tr>
<% if (oTax.FOneItem.Fbilldiv = "01") or (oTax.FOneItem.Fbilldiv = "11") then %>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">��û��</td>
				<td bgcolor="#FFFFFF" colspan="3"><%= oTax.FOneItem.Fuserid %></td>
			</tr>
			<tr height="25">
				<td align="center" bgcolor="#F0F0FD">�Ա�Ȯ����</td>
				<td bgcolor="#FFFFFF" width="120">
					<% if IsNULL(oTax.FOneItem.Fipkumdate) then %>

				 	<% else %>
						<% if oTax.FOneItem.Fipkumdate <> "" then %>
							<%=FormatDate(oTax.FOneItem.Fipkumdate,"0000-00-00")%>
						<% end if %>
			    	<% end if %>
				</td>
				<td align="center" bgcolor="#F0F0FD" width="120">�����</td>
				<td bgcolor="#FFFFFF"><%= oTax.FOneItem.Fregdate %></td>
			</tr>
<% end if %>
<%	if oTax.FOneItem.FisueYn="Y" then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�߱� ����</td>
				<td bgcolor="#F8F8FF" colspan="3">
					<font color=darkblue>�߱�</font>
					&nbsp;
					<% if (Left(oTax.FOneItem.FneoTaxNo,2)="TX") or (Left(oTax.FOneItem.FneoTaxNo,2)="FX") then %>
					<input type="button" class="button" value="���޹޴��� ������" onClick="goView_Bill36524('<%=oTax.FOneItem.FneoTaxNo%>', '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>')" style="cursor:pointer" align="absmiddle">
					<% else %>
					<input type="button" class="button" value="���޹޴��� ������" onClick="goView(<%=oTax.FOneItem.FneoTaxNo%>, '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>', '2118700620')" style="cursor:pointer" align="absmiddle">
					<% end if %>
					<% if (Left(oTax.FOneItem.FneoTaxNo,2)="TX") or (Left(oTax.FOneItem.FneoTaxNo,2)="FX") then %>
					<input type="button" class="button" value="������ ������" onClick="goView_Bill36524('<%=oTax.FOneItem.FneoTaxNo%>', '2118700620')" style="cursor:pointer" align="absmiddle">

					<% else %>
					<input type="button" class="button" value="������ ������" onClick="goView2(<%=oTax.FOneItem.FneoTaxNo%>, '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>', '2118700620')" style="cursor:pointer" align="absmiddle">
				    <% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�߱��� ���̵�</td>
				<td bgcolor="#FFFFFF" colspan="3"><%=oTax.FOneItem.FcurUserId%></td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">��꼭 �ۼ�����</td>
				<td bgcolor="#FFFFFF" colspan="3"><b><%=oTax.FOneItem.FisueDate%></b></td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�߱��Ͻ�</td>
				<td bgcolor="#FFFFFF" colspan="3"><%=oTax.FOneItem.Fprintdate%></td>
			</tr>
<%	else %>
	<% if (orderserial <> "") and (ordercancelyn <> "") then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�ֹ���ȣ</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= orderserial %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�ֹ�����</td>
				<td bgcolor="#FFFFFF" colspan="3">
				    <%= iIpkumDivName %>
				    &nbsp;/&nbsp;
					<% if (ordercancelyn <> "N") then %>
						<font color=red>���</font>
					<% else %>
						����
					<% end if %>
					&nbsp;
					<% IF (ordercancelyn<>"N") then %>
					<strong>[��ҵ� �ֹ����� ���� �Ұ� �մϴ�.]</strong>
					<% elseIF (IpkumDiv<8) then %>
					<strong>[����� �κ���ҵ����� �ݾ� ������ ���� �� ������ ������ ��ǰ��� ���Ŀ� ���� �ϼ���.]</strong>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�����Ѿ�</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber(subtotalprice, 0) %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�ǰ�����</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber((subtotalprice - sumPaymentEtc), 0) %>
					&nbsp;
					<% if (accountdiv = "400") then %>(�޴�������)<% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">��������</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber(sumPaymentEtc, 0) %>
				</td>
			</tr>
	<% end if %>
	<% if (mode <> "new") then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">�߱� ����</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<font color=darkred>�̹߱�</font>
					&nbsp;
					<% if (mode = "view") and (oTax.FOneItem.Fbilldiv = "11") and (oTax.FOneItem.FissueMethod = "eSero") then %>
					<input type="button" class="button" value="����߱޿Ϸ�(eSero)" onClick="TaxEvaleSero()">
					<% else %>
						<% '<input type="button" class="button" value="����(Bill36524 �÷���)" onClick="TaxEvalBill36524()"> %>
						<% '<input type="button" class="button" value="����(Bill36524 API)" onClick="TaxEvalBill36524api()"> %>
						<input type="button" class="button" value="����(���ϰ�)" onClick="PopCommonWehagoTaxReg()">
        				<% '<input type="button" class="button" value="����" onclick="PopCommonSampleTaxReg(); return false;"> %>
					<% end if %>
					&nbsp;
				</td>
			</tr>
	<%	end if %>
<% end if %>
		</table>
	</td>
</tr>
<tr height="20">
	<td>
	</td>
</tr>
<tr height="20">
	<td>
		<br>*<font color=red>�ݾ��� ����</font>�� ��� �ݾ��� ������ �� ����, ������ ���ۼ��ؾ� �մϴ�.
		<br>*����ó�� ��� xx-xxx-xxxx(����) xx.xxx.xxxx(����) �� �Է��� �ּ���
	</td>
</tr>
<tr>
	<td>
		<br>
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="POST" action="doTaxOrder.asp" style="margin:0px;">
			<input type="hidden" name="taxIdx" value="<%= taxIdx %>">
			<input type=hidden name=issuetype value="<%= issuetype %>">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="tplcompanyid" value="<%= tplcompanyid %>">
			<input type="hidden" name="mode" value="">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="15%"><b>�μ�</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="35%">
					<select class="select" name="sellBizCd">
					<option value="">--����--</option>
					<% For i = 0 To UBound(arrBizList,2)	%>
						<option value="<%=arrBizList(0,i)%>" <%IF Cstr(oTax.FOneItem.FsellBizCd) = Cstr(arrBizList(0,i)) THEN%> selected <%END IF%>><%=arrBizList(1,i)%></option>
					<% Next %>
					</select>
    				<%'= fndrawSaleBizSecCombo(true,"sellBizCd", oTax.FOneItem.FsellBizCd,"") %>
    			</td>
    			<td height="25" width="15%"><b>����</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<% drawPartnerCommCodeBox true,"sellacccd","selltype", oTax.FOneItem.Fselltype,"" %>
    			</td>
    		</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25"><b>���γ���</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="taxissuetype">
    					<option value="">-����-</option>
    					<option value="C" <% if (oTax.FOneItem.Ftaxissuetype = "C") then %>selected<% end if %> >�¶����ֹ�</option>
						<option value="F" <% if (oTax.FOneItem.Ftaxissuetype = "F") then %>selected<% end if %> >���������ֹ�</option>
						<option value="E" <% if (oTax.FOneItem.Ftaxissuetype = "E") then %>selected<% end if %> >��Ÿ����</option>
						<!-- <option value="S" <% if (oTax.FOneItem.Ftaxissuetype = "S") then %>selected<% end if %> >�����Ʈ</option> -->
    					<option value="X" <% if (oTax.FOneItem.Ftaxissuetype = "X") then %>selected<% end if %> >��������</option>
    				</select>
    			</td>
    			<td height="25"></td>
    			<td align="left" bgcolor="#FFFFFF">
    			</td>
    		</tr>
    	</table>

<Br>

    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="15%"><b>����Ź����</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="35%">
    				<select class="select" name="consignYN">
    					<option value="">-����-</option>
    					<option value="N" <% if (oTax.FOneItem.FconsignYN = "N") then %>selected<% end if %> >����</option>
						<option value="Y" <% if (oTax.FOneItem.FconsignYN = "Y") then %>selected<% end if %> >����Ź(��ü�Һ��ڸ���)</option>
    				</select>
    			</td>
    			<td height="25" width="15%"><b>��꼭����</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="issueMethod">
    					<option value="">-����-</option>
    					<!--<option value="bill36524" <% 'if (oTax.FOneItem.FissueMethod = "bill36524") then %>selected<% 'end if %> >BILL36524</option>-->
						<option value="WEHAGO" <% if (oTax.FOneItem.FissueMethod = "WEHAGO") then %>selected<% end if %> >���ϰ�</option>
						<option value="eSero" <% if (oTax.FOneItem.FissueMethod = "eSero") then %>selected<% end if %> >�̼��� ����</option>
    				</select>
    			</td>
    		</tr>
    	</table>

<Br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
        <td width="49%">
        	<!-- ���������� ���� -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>������ ����</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
        			<td colspan="3">
        				<input type=text name="reg_socno" size=12 value="<%= oTax.FOneItem.FsupplyBusiNo %>" class="readonlybox" readonly>
        				<select class="select" name="billdiv" onchange="setRegisterInfo()">
							<option value="">�����ڼ���</option>
							<option value="">-----------</option>
							<% if (oTax.FOneItem.Fbilldiv <> "99") then  %>
							<option value="11" <% if (oTax.FOneItem.Fbilldiv = "11") then %>selected<%end if %>>�Һ���(��ü����)</option>
							<option value="01" <% if (oTax.FOneItem.Fbilldiv = "01") then %>selected<%end if %>>�Һ���(2013�����)</option><!-- customer -->
							<option value="">-----------</option>
        					<option value="02" <% if (oTax.FOneItem.Fbilldiv = "02") then %>selected<%end if %>>������(accounts)</option>
        					<option value="03" <% if (oTax.FOneItem.Fbilldiv = "03") then %>selected<%end if %>>���θ��(promotion)</option>
        					<option value="51" <% if (oTax.FOneItem.Fbilldiv = "51") then %>selected<%end if %>>��Ÿ����(accounts)</option>
							<% if (oTax.FOneItem.Fbilldiv = "52") then %>
        						<option value="52" selected>���Ʒ���(youareagirl)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "53") then %>
        						<option value="53" selected>���̾ż�(ithinkso)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "54") then %>
        						<option value="54" selected>���ٸ���(living1010)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "55") then %>
        						<option value="55" selected>�����÷�����(aplusb)</option>
							<%end if %>
							<% else %>
							<option value="99" <% if (oTax.FOneItem.Fbilldiv = "99") then %>selected<%end if %>><%= tplcompanyname %>(<%= tplbillUserID %>)</option>
							<% end if%>
        				</select>
        			</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
        			<td><input type=text name="reg_socname" size=14 value="<%= oTax.FOneItem.FsupplyBusiName %>" border=0 class="readonlybox" readonly></td>
        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
        			<td><input type=text name="reg_ceoname" size=8 value="<%= oTax.FOneItem.FsupplyBusiCEOName %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
        			<td colspan="3"><input type=text name="reg_socaddr" size=40 value="<%= oTax.FOneItem.FsupplyBusiAddr %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">����</td>
        			<td colspan=2><input type=text name="reg_socstatus" size=20 value="<%= oTax.FOneItem.FsupplyBusiType %>" class="readonlybox" readonly></td>
        			<td bgcolor="#F0F0FD">��������ȣ</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">����</td>
        			<td colspan=2><input type=text name="reg_socevent" size=20 value="<%= oTax.FOneItem.FsupplyBusiItem %>" class="readonlybox" readonly></td>
        			<td><input type=text name="reg_subsocno" size=4 value="<%= oTax.FOneItem.FsupplyBusiSubNo %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�����</td>
        			<td><input type=text name="reg_managername" size=14 value="<%= oTax.FOneItem.FsupplyRepName %>" class="readonlybox" readonly></td>
        			<td bgcolor="#F0F0FD">����ó</td>
        			<td><input type=text name="reg_managerphone" size=14 value="<%= oTax.FOneItem.FsupplyRepTel %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
        			<td colspan=3><input type=text name="reg_managermail" size=20 value="<%= oTax.FOneItem.FsupplyRepEmail %>" class="readonlybox" readonly></td>
        		</tr>
        	</table>
        	<!-- ���������� �� -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
			<!-- ���޹޴������� ���� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td colspan="4" height="25" align="right">
						<b>���޹޴��� ����</b>
						&nbsp;
						&nbsp;
						&nbsp;
						<input type="button" class="button" value="[�ŷ�ó�˻�]" onClick="jsGetCust('frm');">
						<input type="button" class="button" value="[����ڰ˻�]" onClick="PopUpcheSelectBySocno('frm');">
						<% '<a href="http://www.nts.go.kr/cal/cal_check_02.asp" target="_blank">[�������ȸ]</a> %>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
					<td colspan="3">
						<input type=text name="socno" size=13 value="<%= oTax.FOneItem.FbusiNo %>" class="writebox">
						<% if (mode = "new") and (issuetype <> "etcmeachul") then %>
						<input type="button" class="button_s" value="�� ��" onClick="SearchSocno()">
						<% end if %>
						<% if (userid <> "") then %>
							<input type="button" class="button_s" value="����(<%= previssuecount %>)" onClick="popListPreviousCustomerTaxSheet('<%= userid %>')" <% if (previssuecount < 1) then %>disabled<% end if %>>
						<% end if %>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
					<td align="left"><input type=text name="socname" size=14 value="<%= oTax.FOneItem.FbusiName %>" border=0 class="writebox"></td>
					<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
					<td align="left"><input type=text name="ceoname" size=14 value="<%= oTax.FOneItem.FbusiCEOName %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">������ּ�</td>
					<td align="left" colspan="3"><input type=text name="socaddr" size=40 value="<%= oTax.FOneItem.FbusiAddr %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">����</td>
					<td colspan=2><input type=text name="socstatus" size=20 value="<%= oTax.FOneItem.FbusiType %>" class="writebox"></td>
					<td bgcolor="#F0F0FD">��������ȣ</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">����</td>
					<td colspan=2><input type=text name="socevent" size=20 value="<%= oTax.FOneItem.FbusiItem %>" class="writebox"></td>
					<td><input type=text name="subsocno" size=4 value="" class="writebox"></td>
				</tr>
				<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">�����</td>
					<td align="left"><input type=text name="managername" size=14 value="<%= oTax.FOneItem.FrepName %>" class="writebox"></td>
					<td bgcolor="#F0F0FD">����ó</td>
					<td align="left"><input type=text name="managerphone" size=14 value="<%= oTax.FOneItem.FrepTel %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">�̸���</td>
					<td align="left" colspan="3"><input type=text name="managermail" size=40 value="<%= oTax.FOneItem.FrepEmail %>" class="writebox"></td>
				</tr>
			</table>
			<!-- ���޹޴������� �� -->
        </td>
	</tr>
</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="120" height="25">������</td>
		<td width="100">���ް���</td>
		<td width="100">��������</td>
		<td width="100">����</td>
		<td width="100">�հ�ݾ�</td>
		<td>���</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25">
			<input type="text" size="10" name="yyyymmdd" value="<%= oTax.FOneItem.FisueDate %>" onClick="jsPopCal('frm','yyyymmdd');" style="cursor:hand;" class="writebox">
		</td>
		<td><input type=text name="totalsupplyprice" size=10 value="" class="readonlybox" readonly></td>
		<td>
			<select name=taxtype class="select" onChange="CalcPrice(this)">
				<option value="">====</option>
				<option value="Y" <% if (oTax.FOneItem.FTaxType = "Y") then %>selected<% end if %>>����</option>
				<option value="N" <% if (oTax.FOneItem.FTaxType = "N") then %>selected<% end if %>>�鼼</option>
				<option value="0" <% if (oTax.FOneItem.FTaxType = "0") then %>selected<% end if %>>����</option>
			</select>
		</td>
		<td><input type=text name="totaltaxprice" size=10 value="<%= (oTax.FOneItem.FtotalTax) %>" class="readonlybox" readonly  onkeyup="CalcPrice(this)"></td>
		<td><input type=text name="totalprice" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="writebox" onkeyup="CalcPrice(this)"></td>
		<% if (mode <> "new") then %>
		<input type="hidden" name="orgtotalprice" value="<%= oTax.FOneItem.FtotalPrice %>">
		<% end if %>
		<td>
			<%
			if (oTax.FOneItem.FtaxIdx = "") then
				'// �ۼ�
				%>
				<input type=text name="orderidx" size=20 value="<%= oTax.FOneItem.Forderidx %>" class="writebox">
				<% if (mode = "new") and (issuetype = "etcmeachul") and (etcmeachulIdx <> "") then %>
				<input type=button class=button name="btnCombine" value="�߰�" onClick="popMeachulDetailList()">
				<% end if %>
				<%
			else
				'// ����
				%>
				<% if (Trim(oTax.FOneItem.Forderserial) <> "") then %>
				�ֹ���ȣ/����ڵ� : <%= oTax.FOneItem.Forderserial %>
				<% elseif (CStr(oTax.FOneItem.Forderidx) <> "0") and (CStr(oTax.FOneItem.Forderidx) <> "") then %>
				�ε����ڵ� : <%=oTax.FOneItem.Forderidx %>
				<% elseif Not IsNull(oTax.FOneItem.GetMultiOrderIdxList) then  %>
				�ε����ڵ� : <%= oTax.FOneItem.GetMultiOrderIdxList %>
				<% end if %>
				<%
			end if
			%>
		</td>
	</tr>
</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="30" height="25">��</td>
		<td width="30">��</td>
		<td>ǰ��</td>
		<td width="50">�԰�</td>
		<td width="50">����</td>
		<td width="100">�ܰ�</td>
		<td width="100">���ް���</td>
		<td width="100">����</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25">
			<%= mid(oTax.FOneItem.FisueDate,6,2) %>
		</td>
		<td><%= mid(oTax.FOneItem.FisueDate,9,2) %></td>
		<td><input type=text name="itemname" size=40 value="<%=db2html(oTax.FOneItem.Fitemname)%>" class="writebox"></td>
		<td></td>
		<td><%= CHKIIF(oTax.FOneItem.Fbilldiv = "01","","1") %></td>

		<td><input type=text name="totalsupplyprice2" size=10 value="" class="writebox" onkeyup="CalcPrice(this)"></td>
		<td><input type=text name="totalsupplyprice3" size=10 value="" class="readonlybox" readonly></td>
		<td><input type=text name="totaltaxprice2" size=10 value="<%= (oTax.FOneItem.FtotalTax) %>" class="readonlybox" readonly></td>
	</tr>
</table>
	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td height="25"><b>�հ�ݾ�</b></td>
		<td width="100">����</td>
		<td width="100">��ǥ</td>
		<td width="100">����</td>
		<td width="100">�ܻ�̼���</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><input type=text name="totalprice2" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly></td>
		<td>
<% if (oTax.FOneItem.Fbilldiv = "01") then %>
			<input type=text name="totalprice3" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly>
<% end if %>
		</td>
		<td></td>
		<td></td>
		<td>
<% if (oTax.FOneItem.Fbilldiv <> "01") then %>
			<input type=text name="totalprice3" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly>
<% end if %>
		</td>
	</tr>
</table>

	</td>
</tr>
<% if (mode = "new") then %>
<tr>
	<td align="right">
		<input type="checkbox" name="ckHand" onClick="chgHandTax(this)">���� �����Է�
	</td>
</tr>
<% end if %>
</form>
<tr>
	<td>

		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		    <tr align="center">
				<td align="center" height="25">
					<% if (mode = "new") then %>
										<input type="button" class="button" value="�ۼ�" onClick="doRegisterSheet()">
					<% else %>
						<% if (oTax.FOneItem.FisueYn = "N") then %>
						<input type="button" class="button" value="����" onClick="GotoTaxModify()">
						&nbsp;
						<% end if %>
						<input type="button" class="button" value="���" onClick="self.location='Tax_list.asp?menupos=<%=menupos %>'">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<% if (oTax.FOneItem.FisueYn = "Y") then %>
					        <% if (oTax.FOneItem.FdelYn = "Y") then %>
					        <input type="button" class="button" value="����(�Ұ�)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" disabled >
					        <font color="red">(������ ������ �����Դϴ�.)</font>
					        <% else %>
					        	<% if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
					        		<input type="button" class="button" value="����(������)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" >
					        	<% else %>
					        		<input type="button" class="button" value="����(�Ұ�)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" disabled >
					        	<% end if %>
					        <% end if %>
						<% ELSE %>
					        <% if (oTax.FOneItem.FdelYn = "Y") then %>
							<input type="button" class="button" value="����(�Ұ�)" disabled >
							(������ ������ �����Դϴ�.)
							<% else %>
							<input type="button" class="button" value="����" onClick="GotoTaxDel('<%= oTax.FOneItem.FisueYn %>')">
							<% end if %>
						<% end if %>
					<% end if %>
				</td>
			</tr>
		</table>

	</td>
</tr>
</table>

<br>

<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<script type='text/javascript'>

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){
	setRegisterInfo();
	CalcPrice(document.frm.taxtype);
}

window.onload = getOnload;

</script>

<%
function Is3PLShopid(shopid)
	dim sqlStr

	Is3PLShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		Is3PLShopid = True
	end if
	rsget.close
end function

function Get3PLUpcheInfoByShopid(shopid, byRef tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		tplcompanyid = rsget("tplcompanyid")
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function

function Get3PLUpcheInfo(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(tplcompanyid) + "' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
