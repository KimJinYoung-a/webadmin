<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ���� ���� ���
' History : 2010.03.22 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->
<%
Dim clsGift , evt_code, cEvent, gift_status, sStateDesc , gift_code ,gift_scope
Dim gift_name, gift_startdate, gift_enddate, opendate, gift_type , gift_itemname
dim makerid , gift_range1 , gift_range2 , giftkind_code , giftkind_type ,giftkind_cnt
dim giftkind_limit, gift_using ,regdate , adminid , giftkind_name , closedate
	gift_code = requestCheckVar(Request("gift_code"),10)

dim itemgubun, shopitemid, itemoption, shopitemname
dim gift_itemgubun, gift_shopitemid, gift_itemoption
dim gift_scope_add, giftkind_limit_sold


set clsGift = new cgift_list
	clsGift.frectgift_code = gift_code
	clsGift.fnGetGiftConts_off

	gift_name			= clsGift.foneitem.fgift_name
	gift_scope 			= clsGift.foneitem.fgift_scope
	evt_code			= clsGift.foneitem.fevt_code
	gift_type			= clsGift.foneitem.fgift_type
	gift_range1			= clsGift.foneitem.fgift_range1
	gift_range2 		= clsGift.foneitem.fgift_range2
	giftkind_code		= clsGift.foneitem.fgiftkind_code

	makerid				= clsGift.foneitem.fmakerid

	itemgubun			= clsGift.foneitem.fitemgubun
	shopitemid			= clsGift.foneitem.fshopitemid
	itemoption			= clsGift.foneitem.fitemoption
	shopitemname		= clsGift.foneitem.fshopitemname

	gift_itemgubun		= clsGift.foneitem.fgift_itemgubun
	gift_shopitemid		= clsGift.foneitem.fgift_shopitemid
	gift_itemoption		= clsGift.foneitem.fgift_itemoption

	giftkind_type		= clsGift.foneitem.fgiftkind_type
	giftkind_cnt		= clsGift.foneitem.fgiftkind_cnt
	giftkind_limit		= clsGift.foneitem.fgiftkind_limit
	gift_startdate		= clsGift.foneitem.fgift_startdate
	gift_enddate		= clsGift.foneitem.fgift_enddate
	gift_status			= clsGift.foneitem.fgift_status
	gift_using     		= clsGift.foneitem.fgift_using
	regdate				= clsGift.foneitem.fregdate
	adminid 			= clsGift.foneitem.fadminid
	giftkind_name 		= clsGift.foneitem.fgiftkind_name
	opendate			= clsGift.foneitem.fopendate
	closedate			= clsGift.foneitem.fclosedate
	gift_itemname		= clsGift.foneitem.fgift_itemname

	'receiptstring		= clsGift.fnGetReceiptString

	gift_scope_add		= clsGift.foneitem.fgift_scope_add
	giftkind_limit_sold	= clsGift.foneitem.fgiftkind_limit_sold

set clsGift = nothing

  '�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
sStateDesc 	= fnSetCommonCodeArr_off("gift_status",False)
%>

<script language="javascript">

var parfrm = parent.opener.document.frmReg;

// XXXXXXXXXXXXXXX �̺�Ʈ�ڵ尡 �ְų�, ����ǰ������ ������ �Ŀ��� ����� �ʴ´�.
// �׻� ����� �ʴ´�. �̺�Ʈ�� ���� ����ǰ�� ����� �� ����.
/*
if (parfrm.gift_name.value == "") {
	parfrm.gift_name.value='<%=gift_name%>';
	parfrm.gift_startdate.value='<%=gift_startdate%>';
	parfrm.gift_enddate.value='<%=gift_enddate%>';
}
*/

parfrm.gift_scope.value = '<%=gift_scope%>';

parfrm.makerid.value = '<%=makerid%>';

parfrm.itemgubun.value = '<%=itemgubun%>';
parfrm.shopitemid.value = '<%=shopitemid%>';
parfrm.itemoption.value = '<%=itemoption%>';
parfrm.shopitemname.value = '<%=shopitemname%>';

parfrm.gift_scope_add.value = '<%=gift_scope_add%>';

parfrm.gift_type.value='<%= gift_type %>';

parfrm.gift_range1.value='<%=gift_range1%>';
parfrm.gift_range2.value='<%=gift_range2%>';

parfrm.gift_itemgubun.value = '<%=gift_itemgubun%>';
parfrm.gift_shopitemid.value = '<%=gift_shopitemid%>';
parfrm.gift_itemoption.value = '<%=gift_itemoption%>';

parfrm.giftkind_code.value='<%=giftkind_code%>';

parfrm.giftkind_name.value = '<%=giftkind_name%>';

parfrm.giftkind_cnt.value = '<%=giftkind_cnt%>';

parent.opener.jsChkGiftScope('<%= gift_scope %>');
parent.opener.jsChkGiftType('<%= gift_type %>');

var igkLmt = '<%=giftkind_limit%>';
if (eval(igkLmt)>0){

	// ������ �����Ǿ� ���� ������ ������ �� �� ����.
	if (parfrm.chkLimit.disabled != true) {
		parfrm.chkLimit.checked=true;
		parfrm.giftkind_limit.value=igkLmt;

		parfrm.giftkind_limit.value = '<%=giftkind_limit%>';

		parent.opener.jsChkLimit();
	}
}

parent.close();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->