<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<%

'// ������
1

'// ��������>>[����]���ݰ�꼭�����û ���� �űԹ����ϸ� ������ ������
'// [OFF]����_���Ͱ���>>�������������(����) ���� �����û �ϸ� ������ ������
'// http://webadmin.10x10.co.kr/cscenter/ordermaster/ordermaster_detail.asp?orderserial=12021576159 ���� �������� �߱� -> ���ݰ�꼭 ���࿡�� ���Ǵ� ������

dim socno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail
dim taxtype, totalpricesum, itemname, totalsuply

dim etcstring, billdiv
dim orderserial, issuetype, idx
dim i, strSql

dim taxwritedate, previssuecount, userid, orderidx
dim errMSG

dim sellBizCd, selltype, taxissuetype

function Is3PLShopid(shopid)
	dim sqlStr

	Is3PLShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "
	''response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		Is3PLShopid = True
	end if
	rsget.close
end function

function Get3PLUpcheInfoByShopid(shopid, byRef tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		tplcompanyid = rsget("tplcompanyid")
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function

orderserial = request("orderserial")
issuetype 	= request("issuetype")
idx 		= request("idx")


itemname = "XXXX �� X ��"

sellBizCd = ""
selltype = "0"

if (orderserial <> "") or (issuetype = "orderserial") then
	'// �Һ��� ����
	etcstring 		= orderserial
	billdiv 		= "01"
	issuetype 		= "orderserial"
	taxtype			= "Y"

	sellBizCd 		= "0000000101"		'// �¶���(����)
	selltype 		= "20166"			'// B2C
	taxissuetype	= "C"

	'==============================================================================
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
		orderidx = ojumun.FOneItem.Fidx
		taxwritedate = getMayTaxDate(ojumun.FOneItem.Fipkumdate)
	end if

	'==============================================================================
	strSql =	"select ( select " &_
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
	rsget.Open strSql, dbget, 1

	if Not(rsget.EOF or rsget.BOF) then
		itemname = rsget("itemname")

		if (CStr(rsget("accountdiv")) = "7") or (CStr(rsget("accountdiv")) = "20") then
			'������, �ǽð���ü : ��ü�ݾ�
			totalpricesum = rsget("subtotalprice")
		else
			'���������ݾ׸�
			totalpricesum = rsget("sumPaymentEtc")
		end if
	end if
	rsget.close

	'==============================================================================
	dim oTax
	set oTax = new CTax

	if (errMSG = "") then
		oTax.FCurrPage = 1
		oTax.FPageSize = 100
		'oTax.FRectsearchDiv = "Y"					'����� ������
		oTax.FRectsearchBilldiv = "01"				'�Һ��ڸ���
		oTax.FRectsearchKey = "t1.userid"
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
end if

dim tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID
dim tplsocno, tplsocname, tplceoname, tplsocaddr, tplsocstatus, tplsocevent, tplmanagername, tplmanagerphone, tplmanagermail

if (issuetype = "etcmeachul") then
	'// ��Ÿ����
	billdiv = "51"
	orderidx = idx

	'==========================================================================
	'��Ÿ����
	dim oetcmeachul
	set oetcmeachul = new CEtcMeachul
	oetcmeachul.FRectidx = idx
	oetcmeachul.getOneEtcMeachul

	sellBizCd 		= oetcmeachul.FOneItem.Fbizsection_cd
	selltype 		= oetcmeachul.FOneItem.Fselltype
	taxissuetype	= "E"

	'oetcmeachul.FOneItem.Ftotalsum '�� ����ݾ��� �Ѱ��ް��� ��.(�ΰ������Աݾ�)


	'==========================================================================
	'����̵𿡼� �׷��ڵ� ����
	dim opartner
	set opartner = new CPartnerUser

	opartner.FCurrpage = 1
	opartner.FPageSize = 100
	opartner.FRectDesignerID = oetcmeachul.FOneItem.Fshopid
	opartner.FRectUserDiv = "all"

	opartner.GetPartnerNUserCList

	'==========================================================================
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
		socno			= ogroup.FOneItem.Fcompany_no
		socname			= ogroup.FOneItem.FCompany_name
		ceoname			= ogroup.FOneItem.Fceoname
		socaddr			= ogroup.FOneItem.Fcompany_address & " " & ogroup.FOneItem.Fcompany_address2
		socstatus		= ogroup.FOneItem.Fcompany_uptae
		socevent		= ogroup.FOneItem.Fcompany_upjong
		managername		= ogroup.FOneItem.Fjungsan_name
		managerphone	= ogroup.FOneItem.Fjungsan_hp
		managermail		= ogroup.FOneItem.Fjungsan_email

		taxtype			= "Y"
		totalpricesum	= oetcmeachul.FOneItem.Ftotalsum
		itemname		= oetcmeachul.FOneItem.Ftitle
		etcstring		= idx
	end if

	'==========================================================================
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
			tplsocno			= otplgroup.FOneItem.Fcompany_no
			tplsocname			= otplgroup.FOneItem.FCompany_name
			tplceoname			= otplgroup.FOneItem.Fceoname
			tplsocaddr			= otplgroup.FOneItem.Fcompany_address & " " & otplgroup.FOneItem.Fcompany_address2
			tplsocstatus		= otplgroup.FOneItem.Fcompany_uptae
			tplsocevent			= otplgroup.FOneItem.Fcompany_upjong
			tplmanagername		= otplgroup.FOneItem.Fjungsan_name
			tplmanagerphone		= otplgroup.FOneItem.Fjungsan_hp
			tplmanagermail		= otplgroup.FOneItem.Fjungsan_email

			billdiv = "99"
		end if
	end if
end if

if (issuetype <> "") and (orderidx <> "") then
	'==========================================================================
	''����� ���ݰ�꼭���� üũ

	set oTax = new CTax

	oTax.FRectsearchKey = " t1.orderidx "
	oTax.FRectsearchString = CStr(orderidx)
	oTax.FRectDelYn = "N"

	oTax.GetTaxList

	if oTax.FResultCount > 0 then
		if oTax.FTaxList(0).FisueYn="Y" then
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

%>
<script language="javascript">
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

// ����ڵ���� Ȯ�� ó��
function chkSheetOk(){
	if (confirm('����ڵ������ Ȯ���ϼ̽��ϱ�?')){
		document.frm_trans.mode.value="BusiOk";
		document.frm_trans.submit();
	}
}

// ��û�� ��� ó��
function GotoTaxPrint(){
    alert('�׿���Ʈ�� ���̻� �������� �ʽ��ϴ�.');
    return;
	if (confirm('���ݰ�꼭�� �����Ͻðڽ��ϱ�?')){
		document.frm_trans.mode.value="sheetOk";
		document.frm_trans.submit();
	}
}

// ��û�� ����
function GotoTaxDel(){
	if (confirm('��û���� ���� �Ͻðڽ��ϱ�?\n\n��꼭�� ����� ��� ������ ��ҵ��� �����Ͻñ� �ٶ��ϴ�.')){
		document.frm_trans.mode.value="sheetDel";
		document.frm_trans.submit();
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

function setRegisterInfo()
{
	var isUnitTax, isUnitTaxTypeChanged;

	// 2012-01-01 ���� �������� ���ݰ�꼭�� �����Ѵ�.
	if (frm.yyyymmdd_register.value == "") {
		// �ۼ����� ������ �������� ���ݰ�꼭
		frm.prev_yyyymmdd_register.value = "2012-01-01";
		isUnitTax = true;
		isUnitTaxTypeChanged = false;
	} else {
		if (frm.yyyymmdd_register.value >= "2012-01-01") {
			isUnitTax = true;
		} else {
			isUnitTax = false;
		}

		if (((frm.yyyymmdd_register.value >= "2012-01-01") && (frm.prev_yyyymmdd_register.value >= "2012-01-01")) || ((frm.yyyymmdd_register.value < "2012-01-01") && (frm.prev_yyyymmdd_register.value < "2012-01-01"))) {
			isUnitTaxTypeChanged = false;
		} else {
			isUnitTaxTypeChanged = true;
		}

		frm.prev_yyyymmdd_register.value = frm.yyyymmdd_register.value;
	}

	if (isUnitTaxTypeChanged == true) {
		alert("�ۼ����� ����Ǿ����ϴ�. 2012-01-01 ���� �������� ���ݰ�꼭�� �����մϴ�.");
	}

	// ================================================================
	// cs_taxsheetcls.asp ���� �����´�.
	// ================================================================
	frm.reg_socno.value = "<%= TENBYTEN_SOCNO %>";
	frm.reg_subsocno.value = "<%= TENBYTEN_SUBSOCNO %>";
	frm.reg_socname.value = "<%= TENBYTEN_SOCNAME %>";
	frm.reg_ceoname.value = "<%= TENBYTEN_CEONAME %>";
	frm.reg_socaddr.value = "<%= TENBYTEN_SOCADDR %>";
	frm.reg_socstatus.value = "<%= TENBYTEN_SOCSTATUS %>";
	frm.reg_socevent.value = "<%= TENBYTEN_SOCEVENT %>";
	frm.reg_managername.value = "<%= TENBYTEN_MANAGERNAME %>";
	frm.reg_managerphone.value = "<%= TENBYTEN_MANAGERPHONE %>";
	frm.reg_managermail.value = "<%= TENBYTEN_MANAGERMAIL %>";

	// ===================================================================
	if(frm.billdiv.value == "52") {
		// ������ (��)���ش�����
        alert('����Ұ�');
        return;

		frm.reg_socno.value = "101-85-29011";
		frm.reg_socname.value = "(��)���ش�����";
		frm.reg_ceoname.value = "�̹���";
		frm.reg_socaddr.value = "���� ���α� ��ȭ�� 197-1 �̿������� 2��";
		frm.reg_socstatus.value = "����,���Ҹ�";
		frm.reg_socevent.value = "���ڻ�ŷ� ��";
		frm.reg_managername.value = "����";
		frm.reg_managerphone.value = "02-554-2033";
		frm.reg_managermail.value = "accounts@10x10.co.kr";
	}

	if(frm.billdiv.value == "55") {
		// ������ (��)�����÷�����
        alert('����Ұ�');
        return;

		frm.reg_socno.value = "101-86-64617";
		frm.reg_socname.value = "(��)�����÷�����";
		frm.reg_ceoname.value = "��â��";
		frm.reg_socaddr.value = "���� ���α� ��ȭ�� 197-1 �̿�������2��";
		frm.reg_socstatus.value = "���Ҹ�";
		frm.reg_socevent.value = "���ڻ�ŷ�";
		frm.reg_managername.value = "���ȯ";
		frm.reg_managerphone.value = "070-7515-5410"
		frm.reg_managermail.value = "gogo27@10x10.co.kr"
	}

	if (isUnitTax == true) {
		if(frm.billdiv.value == "53") {
			// ������ (��)���̶��

			frm.reg_subsocno.value = "0001";
		}

		if(frm.billdiv.value == "54") {
			// ������ (��)�ٹ����� ����

			frm.reg_subsocno.value = "0002";
		}

		if(frm.billdiv.value == "99") {
			// ������ 3PL��ü

			frm.reg_socno.value = "<%= tplsocno %>";
			frm.reg_socname.value = "<%= tplsocname %>";
			frm.reg_ceoname.value = "<%= tplceoname %>";
			frm.reg_socaddr.value = "<%= tplsocaddr %>";
			frm.reg_socstatus.value = "<%= tplsocstatus %>";
			frm.reg_socevent.value = "<%= tplsocevent %>";
			frm.reg_managername.value = "<%= tplmanagername %>";
			frm.reg_managerphone.value = "<%= tplmanagerphone %>";
			frm.reg_managermail.value = "<%= tplmanagermail %>";
		}
	} else {
		if(frm.billdiv.value == "53") {
			// ������ (��)���̶��

			frm.reg_socno.value = "101-85-36109";
			frm.reg_socname.value = "(��)���̶��";
			frm.reg_ceoname.value = "�̹���";
			frm.reg_socaddr.value = "���� ���α� ������ 1-45 �������� 4��";
			frm.reg_socstatus.value = "���Ҹ�";
			frm.reg_socevent.value = "�ҽÿ�ǰ";
			frm.reg_managername.value = "���ȯ";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}

		if(frm.billdiv.value == "54") {
			// ������ (��)�ٹ����� ����

			frm.reg_socno.value = "101-85-38408";
			frm.reg_socname.value = "(��)�ٹ����� ����";
			frm.reg_ceoname.value = "�̹���";
			frm.reg_socaddr.value = "���� ���α� ������ 1-45 �������� 1��";
			frm.reg_socstatus.value = "���Ҹ�";
			frm.reg_socevent.value = "�ķ�ǰ,�ҽÿ�ǰ";
			frm.reg_managername.value = "���ȯ";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}
	}
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
	// location.href="isearchframe.asp?socno=" + frm.socno.value;
}

function popMeachulDetailList() {
	if (frm.taxissuetype.value != "E") {
		alert("��Ÿ������ ��쿡�� ������ �߰��� �� �ֽ��ϴ�.");
		return;
	}

	var popwin = window.open('pop_etc_meachul_list.asp?idx=<%= etcstring %>','popMeachulDetailList','width=1000, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
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

function CalcPrice()
{
	if (frm.totalsuply.value == "") { return; }

	if (frm.taxtype.value.length<1){
		alert('���������� �Է��ϼ���.');
		return;
	}

	if (frm.totalsuply.value*0 != 0) { alert("�߸��� ���� �Է��߽��ϴ�."); return; }

	frm.totalsuply2.value = frm.totalsuply.value;
	frm.totalsuplysum.value = frm.totalsuply.value;

    if (frm.ckHand.checked){
        frm.totaltaxsum.value = frm.totaltax.value;
    }else{
		if (frm.taxtype.value == "Y") {
			frm.totaltax.value = parseInt(frm.totalsuply.value*0.1);
			frm.totaltaxsum.value = parseInt(frm.totalsuply.value*0.1);
		} else {
			frm.totaltax.value = 0;
			frm.totaltaxsum.value = 0;
		}
	}

	frm.totalpricesum.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
	frm.totalpricesum2.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
	frm.totalpricesum3.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
}

function CalcPriceWithPrice()
{
	if (frm.totalpricesum.value == "") { return; }

	if (frm.taxtype.value.length<1){
		alert('���������� �Է��ϼ���.');
		return;
	}

	if (frm.totalpricesum.value*0 != 0) { alert("�߸��� ���� �Է��߽��ϴ�."); return; }

	frm.totalpricesum2.value = frm.totalpricesum.value;
	frm.totalpricesum3.value = frm.totalpricesum.value;

	if (frm.taxtype.value == "Y") {
		// ������ ���ް��� ���ϰ� 0.1 �� �ݿø� ���ָ� �ȴ�.
		frm.totaltax.value = Math.round(1.0 * frm.totalpricesum.value / 1.1 / 10.0);
		frm.totaltaxsum.value = frm.totaltax.value;
	} else {
		frm.totaltax.value = 0;
		frm.totaltaxsum.value = 0;
	}

	frm.totalsuply.value = frm.totalpricesum.value - frm.totaltax.value;
	frm.totalsuply2.value = frm.totalsuply.value;
	frm.totalsuplysum.value = frm.totalsuply.value;
}


function doRegisterSheet(){

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if (frm.issuetype.value != "") {
		if ((frm.issuetype.value == "orderserial") && (frm.billdiv.value != "01")) {
			alert('�Һ��� ���⸸ �ۼ��� �� �ֽ��ϴ�.');
			frm.billdiv.focus();
			return;
		}

		if ((frm.issuetype.value == "etcmeachul") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
			alert('��Ÿ���⸸ �ۼ��� �� �ֽ��ϴ�.');
			frm.billdiv.focus();
			return;
		}

		if(frm.etcstring.value == "") {
			alert('�ֹ���ȣ �Ǵ� ��Ÿ���� �ڵ尡 ��� �ԷµǾ� �־�� �մϴ�.');
			return;
		}
	}

	if ((frm.selltype.value == "20036") && (frm.taxtype.value != "0")) {
		alert('���������� ������ ��� ������꼭�� �ۼ������մϴ�.');
		return;
	} else if ((frm.selltype.value != "20036") && (frm.taxtype.value == "0")) {
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

	if (frm.managermail.value.length<1){
		alert('����� �̸����ּҸ� �Է��ϼ���.');
		frm.managermail.focus();
		return;
	}

	if (frm.yyyymmdd_register.value.length<1){
		alert('�ۼ����� �Է��ϼ���.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('ǰ���� �Է��ϼ���.');
		return;
	}

	if (frm.totalsuply.value.length<1){
		alert('�ܰ��� �Է��ϼ���.');
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

	if(frm.billdiv.value == "01") {
		if(frm.etcstring.value == "") {
			alert('��� �ֹ���ȣ �Ǵ� ����ڵ带 �Է��ϼ���.');
			return;
		}
	} else if ((frm.etcstring.value != "") && (frm.billdiv.value != "03") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
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

	setRegisterInfo();

    if (confirm('���ݰ�꼭 �����û�� �Ͻðڽ��ϱ�?')){
        document.frm.submit();
    }
}

function chgHandTax(comp){
    var txbox = comp.form.totaltax;

    if (comp.checked){
        txbox.readOnly = false;
        txbox.className = "writebox";
    }else{
        txbox.readOnly = true;
        txbox.className = "readonlybox";
    }
}

function popListPreviousCustomerTaxSheet(userid){
    var popwin=window.open("/cscenter/taxsheet/popListPreviousCustomerTaxSheet.asp?userid=" + userid,"popListPreviousCustomerTaxSheet","width=700,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ReactMeachulDetailList(arrchk, tottaxsum) {
    var frm = document.frm;

    frm.totalpricesum.value = tottaxsum;
    frm.etcstring.value = arrchk;

    CalcPriceWithPrice();
}

</script>

<style type="text/css">
.readonlybox { border:0px; }
.writebox { border:0px; background:#E6E6E6; }
</style>



<table width="800" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td colspan="2" align="left">
					<b>���ݰ�꼭 �����û</b>
				</td>
			</tr>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">��û��</td>
				<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
			</tr>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">����</td>
				<td bgcolor="#FFFFFF">
					�����ڴ� <b>�ٹ�����/���̶��/���Ʒ���</b> �߿� �ϳ��� �����Ͻø�, �ڵ��Էµ˴ϴ�.<br>
					���޹޴��ڴ� ��Ϲ�ȣ�� ������(-)�� ������ ����ڹ�ȣ�� �Է��Ͻø�, ���� ����Ÿ�� ������� �ڵ��Էµ˴ϴ�.<br>
					<b>(�˻� ��, ����ڵ��� ������ ������ ���, �����Է��Ͻø� �˴ϴ�.)</b><br>
					ǰ���� �� �Է��Ͻñ� �ٶ��ϴ�.(���� ��ǰ������ �Ѿ����θ� �Է� �����մϴ�.)<br>
				</td>
			</tr>
		</table>

	</td>
</tr>
<tr height="70">
	<td>
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" onsubmit="return false;" action="doTaxOrder.asp">
			<input type=hidden name=mode value="tax_register_new">
			<input type=hidden name=issuetype value="<%= issuetype %>">
			<input type=hidden name=tplcompanyid value="<%= tplcompanyid %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="10%"><b>�μ�</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="40%">
    				<%= fndrawSaleBizSecCombo(true,"sellBizCd", sellBizCd,"") %>
    			</td>
    			<td height="25" width="10%"><b>����</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<% drawPartnerCommCodeBox true,"sellacccd","selltype", selltype,"" %>
    			</td>
    		</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25"><b>���γ���</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="taxissuetype">
    					<option value="">-����-</option>
    					<option value="C" <% if (taxissuetype = "C") then %>selected<% end if %> >�¶����ֹ�</option>
    					<option value="E" <% if (taxissuetype = "E") then %>selected<% end if %> >��Ÿ����</option>
    					<option value="S" <% if (taxissuetype = "S") then %>selected<% end if %> >�����Ʈ</option>
    					<option value="X" <% if (taxissuetype = "X") then %>selected<% end if %> >��������</option>
    				</select>
    			</td>
    			<td height="25"></td>
    			<td align="left" bgcolor="#FFFFFF">
    			</td>
    		</tr>
    	</table>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			<tr valign="top">
		        <td width="49%">
		        	<!-- ���������� ���� -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="2" height="25"><b>������ ����</b></td>
		        			<td colspan="2">
		        				<select class="select" name="billdiv" onchange="setRegisterInfo()">
		        					<option value="0">�����ڼ���</option>
									<% if (billdiv <> "99") then %>
		        					<option value="01" <% if (billdiv = "01") then %>selected<% end if %>>�Һ���(customer)</option>
		        					<option value="02" <% if (billdiv = "02") then %>selected<% end if %>>������(accounts)</option>
		        					<option value="03" <% if (billdiv = "03") then %>selected<% end if %>>���θ��(promotion)</option>
		        					<option value="51" <% if (billdiv = "51") then %>selected<% end if %>>��Ÿ����(accounts)</option>
		        					<!-- option value="52">���Ʒ���(youareagirl)</option -->
									<!--
		        					<option value="53" <% if (billdiv = "53") then %>selected<% end if %>>���̶��(ithinkso)</option>
									-->
		        					<option value="54" <% if (billdiv = "54") then %>selected<% end if %>>�ٹ����� ����(living1010)</option>
		        					<!-- <option value="55">�����÷�����(aplusb)</option> -->
									<% else %>
									<option value="99" <% if (billdiv = "99") then %>selected<% end if %>><%= tplcompanyname %>(<%= tplbillUserID %>)</option>
									<% end if %>
		        				</select>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
		        			<td colspan="3">
		        				<input type=text name="reg_socno" size=12 value="" class="readonlybox" readonly>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
		        			<td><input type=text name="reg_socname" size=14 value="" border=0 class="readonlybox" readonly></td>
		        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
		        			<td><input type=text name="reg_ceoname" size=8 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
		        			<td colspan="3"><input type=text name="reg_socaddr" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td colspan=2><input type=text name="reg_socstatus" size=20 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">��������ȣ</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td colspan=2><input type=text name="reg_socevent" size=20 value="" class="readonlybox" readonly></td>
		        			<td><input type=text name="reg_subsocno" size=4 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�����</td>
		        			<td><input type=text name="reg_managername" size=14 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">����ó</td>
		        			<td><input type=text name="reg_managerphone" size=14 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
		        			<td colspan=3><input type=text name="reg_managermail" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        	</table>
		        	<!-- ���������� �� -->
		        </td>
		        <td>&nbsp;</td>
		        <td width="49%">
		        	<!-- ���޹޴������� ���� -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="4" height="25"><b>���޹޴��� ����</b></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
		        			<td colspan="3">
		        				<input type=text name="socno" size=12 value="<%= socno %>" class="writebox">
		        				<input type="button" class="button_s" value="�� ��" onClick="SearchSocno()">
		        				<% if (userid <> "") then %>
		        					<input type="button" class="button_s" value="����(<%= previssuecount %>)" onClick="popListPreviousCustomerTaxSheet('<%= userid %>')" <% if (previssuecount < 1) then %>disabled<% end if %>>
		        				<% end if %>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
		        			<td align="left"><input type=text name="socname" size=14 value="<%= socname %>" border=0 class="writebox"></td>
		        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
		        			<td align="left"><input type=text name="ceoname" size=14 value="<%= ceoname %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
		        			<td align="left" colspan="3"><input type=text name="socaddr" size=40 value="<%= socaddr %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td colspan=2><input type=text name="socstatus" size=20 value="<%= socstatus %>" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">��������ȣ</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td colspan=2><input type=text name="socevent" size=20 value="<%= socevent %>" class="writebox"></td>
		        			<td><input type=text name="subsocno" size=4 value="" class="writebox"></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�����</td>
		        			<td align="left"><input type=text name="managername" size=14 value="<%= managername %>" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">����ó</td>
		        			<td align="left"><input type=text name="managerphone" size=14 value="<%= managerphone %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
		        			<td align="left" colspan="3"><input type=text name="managermail" size=40 value="<%= managermail %>" class="writebox"></td>
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
				<td width="120" height="25">�ۼ���</td>
				<td width="100">���ް���</td>
				<td width="100">��������</td>
				<td width="100">����</td>
				<td width="100">�հ�ݾ�</td>
				<td>���</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25">
					<input type="text" size="10" name="yyyymmdd_register" value="<%= taxwritedate %>" onClick="jsPopCal('frm','yyyymmdd_register');" style="cursor:hand;" class="writebox">
					<input type=hidden name=prev_yyyymmdd_register value="<%= taxwritedate %>">
				</td>
				<td><input type=text name="totalsuplysum" size=10 value="" class="readonlybox" readonly></td>
				<td>
					<select name=taxtype class="select" onchange="CalcPriceWithPrice()">
					<option value="">====</option>
					<option value="Y" <% if (taxtype = "Y") then %>selected<% end if %>>����</option>
					<option value="N" <% if (taxtype = "N") then %>selected<% end if %>>�鼼</option>
					<option value="0" <% if (taxtype = "0") then %>selected<% end if %>>����</option>
					</select>
				</td>
				<td><input type=text name="totaltaxsum" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totalpricesum" size=10 value="<%= totalpricesum %>" class="writebox" onkeyup="CalcPriceWithPrice()"></td>
				<td>
					<input type=text name="etcstring" size=20 value="<%= etcstring %>" class="writebox">
					<input type=button class=button name="btnCombine" value="�߰�" onClick="popMeachulDetailList()">
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
				<td height="25"></td>
				<td></td>
				<td><input type=text name="itemname" size=40 value="<%= itemname %>" class="writebox"></td>
				<td></td>
				<td>1</td>
				<td><input type=text name="totalsuply" size=10 value="" class="writebox" onkeyup="CalcPrice()"></td>
				<td><input type=text name="totalsuply2" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totaltax" size=10 value="" class="readonlybox" readonly  onKeyUp="CalcPrice();"></td>
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
				<td height="25"><input type=text name="totalpricesum3" size=10 value="" class="readonlybox" readonly></td>
				<td>
				</td>
				<td></td>
				<td></td>
				<td>
					<input type=text name="totalpricesum2" size=10 value="" class="readonlybox" readonly>
				</td>

			</tr>

		</table>

	</td>
</tr>
<tr height="5">
	<td align="right">
		<input type="checkbox" name="ckHand" onClick="chgHandTax(this)">���� �����Է�
	</td>
</tr>
</form>
<tr>
	<td>

		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		    <tr align="center">
				<td align="center" height="25">
				  <input type="button" class="button" value="�ۼ�" onClick="doRegisterSheet()">
				  &nbsp;
				  <input type="button" class="button" value="���" onClick="self.location='Tax_list.asp'">
				</td>
			</tr>
		</table>

	</td>
</tr>
</table>

<p>

<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<script>

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){

	if (frm.billdiv.value != "0") {
		setRegisterInfo();
		CalcPriceWithPrice();
	}

	if (errMSG != "") {
		alert(errMSG);
	}
}

window.onload = getOnload;

</script>


<!-- ���ݰ�� ��û�� ���� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
