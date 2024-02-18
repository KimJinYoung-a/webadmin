<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [LOG]��������>>�����Ʈ
' History : �̻� ����
'			2017.05.26 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
'##################################################################################
' 	������ 		ǰ�ǻ���		�����Ʈ ��������
'   0 -�ۼ��� 					 	��������
'    1- �ֹ�����					��������
'    1- �ֹ�����	ǰ��������		�����Ұ���
'    1- �ֹ�����	ǰ�ǿϷ�		�����Ұ���
'    7-���Ϸ�		ǰ�ǿϷ�		�����Ұ��� (��������ڸ� ���Ϸ� ó�� ����)
'##################################################################################

dim oipchul, oipchuldetail,executedt, storeid,storemarginrate, idx, sqlStr, sellcashtotal, suplycashtotal, buycashtotal, ttlitemno, i
dim ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv, company_no, ischulgonotdisp
dim deldt, currencyUnit_Pos
	idx = request("idx")
	storemarginrate = request("storemarginrate")

ischulgonotdisp=false

set oipchul = new CIpChulStorage
	oipchul.FRectId = idx
	oipchul.GetIpChulMaster

executedt = oipchul.FOneItem.Fexecutedt
deldt = oipchul.FOneItem.Fdeldt

if (Left(oipchul.FOneItem.Fcode,2) <> "SO") then
	response.write "<script>alert('���� : ����ڵ尡 �ƴմϴ�.');</script>"
	response.write "<br><br>���� : ����ڵ尡 �ƴմϴ�." & oipchul.FOneItem.Fcode
	response.end
end if

if (deldt <> "" or not isNull(deldt)) then
	response.write "<script>alert('���� : ������ �����Դϴ�.');</script>"
	response.write "<br><br>���� : ������ �����Դϴ�."
	response.end
end if

set oipchuldetail = new CIpChulStorage
	oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
	oipchuldetail.GetIpChulDetail

sellcashtotal  = 0
suplycashtotal = 0
buycashtotal = 0

dim BasicMonth, IsExpireEdit
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

if IsNULL(oipchul.FOneItem.Fexecutedt) then
	IsExpireEdit = Lcase(CStr(false))
else
	IsExpireEdit = Lcase(CStr(CDate(oipchul.FOneItem.Fexecutedt)<Cdate(BasicMonth)))
end if

if (  (storemarginrate = "") or (storemarginrate = "0") ) then
	sqlStr = "select IsNull(a.marginrate, 0) as marginrate "
	sqlStr = sqlStr + " from [db_storage].[dbo].vw_acount_user_delivery a "
	sqlStr = sqlStr + " where a.userid = '" +  oipchul.FOneItem.Fsocid  + "' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		storemarginrate = rsget("marginrate")
	else
		storemarginrate = "0"
	end if
	rsget.close
elseif (storemarginrate = "") then
	storemarginrate = "0"
end if

if oipchul.FOneItem.Fsocid<>"" then
	ArrShopInfo = getoffshopuser(oipchul.FOneItem.Fsocid)

	IF isArray(ArrShopInfo) then
		currencyunit = ArrShopInfo(1,0)
		currencyChar = ArrShopInfo(3,0)
		loginsite = ArrShopInfo(2,0)
		shopdiv = ArrShopInfo(12,0)
    END IF

	sqlStr = "select p.id, p.company_no, s.currencyUnit, s.currencyUnit_Pos" & vbcrlf
	sqlStr = sqlStr & " from db_partner.dbo.tbl_partner p" & vbcrlf
	sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user s" & vbcrlf
	sqlStr = sqlStr & " 	on p.id=s.userid" & vbcrlf
	sqlStr = sqlStr & " where p.id = '"& oipchul.FOneItem.Fsocid &"'" & vbcrlf

    'response.write sqlStr & "<br>"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		company_no = rsget("company_no")
		currencyUnit = rsget("currencyUnit")
		currencyUnit_Pos = rsget("currencyUnit_Pos")
	end if
	rsget.close
end if

' ���ų� �ؿ��ϰ�� �ٹ����� ����ڰ� �ƴҰ�� �̸Ŵ����� ������.
if IsNull(company_no) then
	company_no = ""
end if

if C_ADMIN_AUTH or C_AUTH or C_MngPart then
else
	if replace(company_no,"-","")<>"2118700620" and (shopdiv="5" or shopdiv="7") then
		ischulgonotdisp = true
	end if
end if

dim ochulgolog
set ochulgolog = new CIpChulStorage
	ochulgolog.FRectStoragecode = oipchul.FOneItem.Fcode
	ochulgolog.FPageSize = 50
	ochulgolog.FCurrPage = 1
	ochulgolog.GetIpChulDetail_edit_log

%>
<script language='javascript'>
function ConvertWiChulgo(iid){
	var popwin = window.open('chulgoedit_process.asp?mode=wichulgoconv&id=' + iid ,'chulgodetailconv','width=300,height=300,scrollbars=yes,resizable=no');
	popwin.focus();
}

function checkAvail2(monthdiff,orgdate){
	var nowdate = "<%= Left(now(),10) %>";
	var odate1 = new Date(orgdate.substring(0,4)*1,orgdate.substring(5,7)*1-1,orgdate.substring(8,10),0,0,0);
	var odate2 = new Date(nowdate.substring(0,4)*1,nowdate.substring(5,7)*1-1-(1-1*monthdiff),0,0,0,0);
	//alert(odate1);
	//alert(odate2);
	if (odate2>=odate1){
		//alert('T');
		return false;
	}else{
		return true;
	}
}

function checkAvail(diffdate,orgdate){
	var nowdate = "<%= Left(now(),10) %>";
	var odate1 = new Date(orgdate.substring(0,4)*1,orgdate.substring(5,7)*1-1,orgdate.substring(8,10),0,0,0);
	var odate2 = new Date(nowdate.substring(0,4)*1,nowdate.substring(5,7)*1-1,nowdate.substring(8,10)-diffdate*1,0,0,0);

	if (odate2>odate1){
		//alert('T');
		return false;
	}else{
		return true;
	}
}

var orgexecutedt = "<%=executedt%>";
function ModiMaster(frm){
<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
	if (<%= IsExpireEdit %>){
		alert('�δ� ���� ���� ������ ���� �Ұ����մϴ�.');
		return;
	}

	if ((orgexecutedt!='')&&(frm.executedt.value<'<%= BasicMonth %>')){
		alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
		return;
	}
<% end if %>

	if (checkAvail3(frm.executedt.value) != true) {
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?')){
		frm.action = "/admin/newstorage/chulgoedit_process.asp";
		frm.mode.value = "editmaster";
		frm.submit();
	}
}

function DelMaster(frm){
<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
	if (<%= IsExpireEdit %>){
		alert('�δ� ���� ���� ������ ���� �Ұ����մϴ�.');
		return;
	}

	if ((orgexecutedt!='')&&(frm.executedt.value<'<%= BasicMonth %>')){
		alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
		return;
	}
<% end if %>

	if (checkAvail3(frm.executedt.value) != true) {
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?')){
		frm.action = "/admin/newstorage/chulgoedit_process.asp";
		frm.mode.value = "delete";
		frm.submit();
	}
}



  function SubmitForm()
  {

          if (document.frmmaster.storeid.value == "") {
                  alert("���ó�� �����ϼ���.");
                  return;
          }
          if (document.frmmaster.divcode.value == "") {
                  alert("������� �����ϼ���.");
                  return;
          }
          if (document.frmmaster.vatcode.value == "") {
                  alert("�ΰ��������� �����ϼ���.");
                  return;
          }
          if (document.frmmaster.scheduledt.value == "") {
                  alert("����û���� �Է��ϼ���.");
                  return;
          }

           if (confirm("�����Ͻðڽ��ϱ�?") != true) {
                  return;
	  		}

          document.frmmaster.mode.value = "write";
          document.frmmaster.action = "chulgoedit_process.asp";
          document.frmmaster.submit();

  }

  	function tempSave(){
		  if (document.frmmaster.storeid.value == "") {
                  alert("���ó�� �����ϼ���.");
                  return;
          }

		 document.frmmaster.mode.value = "temp";
         document.frmmaster.action = "chulgoedit_process.asp";
         document.frmmaster.submit();
	}

// �ſ� 3�ϱ��� �������� ��������
function checkAvail3(modiexecutedt) {
	var orgexecutedt = "<%=executedt%>";
	var thisDate = "<%= Left(Now, 7) %>-01";
	var availDate = "<%= Left(Now, 7) %>-03";
	var nowdate = "<%= Left(now(),10) %>";
	var BasicMonth = "<%= BasicMonth %>";

	if ((orgexecutedt == "") && (modiexecutedt == "")) {
		// ������� ������ ��ŵ
		return true;
	}

	if ((orgexecutedt < BasicMonth) || (modiexecutedt < BasicMonth)) {
		<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
		alert('����Ұ�\n\n������� �δ� ������¥�Դϴ�.');
		return false;
		<% else %>
		alert('�����ڱ���\n\n������� �δ� ������¥�Դϴ�.');
		//alert(orgexecutedt + ' ' + BasicMonth);
		//alert(modiexecutedt + ' ' + BasicMonth);
		<% end if %>
	}

	//alert(thisDate + "/" + orgexecutedt + '/' + modiexecutedt + '/' + availDate)
	//������� �̹��� ���� �������
	if ('<%= Left(Now, 7) %>' > modiexecutedt){
		if ((orgexecutedt < thisDate) || (modiexecutedt < thisDate)) {
			if (nowdate > availDate) {
				<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
					alert("����Ұ�\n\n�ſ� 3�ϱ����� �������� ���氡���մϴ�.");
					return false;
				<% else %>
					alert('�����ڱ���\n\n�ſ� 3�ϱ����� �������� ���氡���մϴ�.');
				<% end if %>
			}
		}
	}

	return true;
}

function ChulgoMaster(frm){
    <% if (oipchul.FOneItem.Fexecutedt <> "") then %>
		alert("�̹� ���ó�� �Ͽ����ϴ�.");
		return;
	<% end if %>

	if (frm.executedt.value.length<1){
		alert('������� �Է��ϼ���.');
		calendarOpen(frm.executedt);
		return;
	}

	<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
		if (frm.executedt.value>'<%= date() %>'){
			alert('������� ���ó�¥ ���� Ŭ�� �����ϴ�.');
			return;
		}
		if ((orgexecutedt!='')&&(frm.executedt.value<'<%= BasicMonth %>')){
			alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
			return;
		}
	<% end if %>

	if (confirm('��� ó�� �Ͻðڽ��ϱ�?')){
		frm.action = "/admin/newstorage/chulgoedit_process.asp";
		frm.mode.value = "chulgo";
		frm.submit();
	}
}

function Chulgo2Jupsu(frm) {
    <% if IsNull(oipchul.FOneItem.Fexecutedt) then %>
	alert("������� �����Դϴ�.");
	return;
	<% elseif Month(oipchul.FOneItem.Fexecutedt) <> Month(Now()) then %>
	alert("����� ������ ������� ��ȯ �����մϴ�.");
	return;
	<% end if %>

	if (confirm('������� ��ȯ �Ͻðڽ��ϱ�?')){
		frm.action = "/admin/newstorage/chulgoedit_process.asp";
		frm.mode.value = "chulgo2jupsu";
		frm.submit();
	}
}

<% if (C_ADMIN_AUTH or C_AUTH) and (oipchul.FOneItem.Fexecutedt <> "") then %>
// ������� ����
function ChChulgoDate(frm) {
	if (frm.executedt.value.substring(0,7) < '<%= Left(oipchul.FOneItem.Fexecutedt,7) %>') {
		if (confirm("������ ������ڸ� ������ ���, ������ ���� ���ۼ��Ͼ� �մϴ�.\n\n�����Ͻðڽ��ϱ�?") !== true) {
			return;
		}
	}

	if (confirm('������� ����\n\n- ���ݿ�\n- �����԰� �������(������ ��ո��԰� �ƴ�)\n\n������� ���� �Ͻðڽ��ϱ�?')){
		frm.action = "/admin/newstorage/chulgoedit_process.asp";
		frm.mode.value = "chchulgodate";
		frm.submit();
	}
}
<% end if %>

function DelDetail(masterfrm,iid){
<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
	if (<%= IsExpireEdit %>){
		alert('�δ� ���� ���� ������ ���� �Ұ����մϴ�.');
		return;
	}

	if ((orgexecutedt!='')&&(masterfrm.executedt.value<'<%= BasicMonth %>')){
		alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
		return;
	}
<% end if %>

	if (checkAvail3(masterfrm.executedt.value) != true) {
		return;
	}

	var frm;
	var found = false;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "chk") {
			if (frm.checked == true) {
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("���õ� ��ǰ�� �����մϴ�.") == true) {
			frmDetail.mode.value = "deldetail";
			frmDetail.action = "/admin/newstorage/chulgoedit_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("���õ� ��ǰ�� �����ϴ�.");
	}
}

function ModiDetail(masterfrm,iid){
<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
	if (<%= IsExpireEdit %>){
		alert('�δ� ���� ���� ������ ���� �Ұ����մϴ�.');
		return;
	}

	if ((orgexecutedt!='')&&(masterfrm.executedt.value<'<%= BasicMonth %>')){
		alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
		return;
	}
<% end if %>

	if (checkAvail3(masterfrm.executedt.value) != true) {
		return;
	}

	var frm;
	var found = false;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "chk") {
			if (frm.checked == true) {
				// elements[i+4]:itemno, elements[i+5]:sellcash, elements[i+6]:suplycash, elements[i+7]:buycash
				if (((frmDetail.elements[i+4].value*0) != 0) || ((frmDetail.elements[i+5].value*0) != 0) || ((frmDetail.elements[i+6].value*0) != 0) || ((frmDetail.elements[i+7].value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return false;
				}
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("���õ� ��ǰ�� �����մϴ�.") == true) {
			frmDetail.mode.value = "editdetail";
			frmDetail.action = "/admin/newstorage/chulgoedit_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("���õ� ��ǰ�� �����ϴ�.");
	}

}

function regAGVArr(){
	var frm;
	var found = false;

	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "chk") {
			if (frm.checked == true) {
				// elements[i+4]:itemno, elements[i+5]:sellcash, elements[i+6]:suplycash, elements[i+7]:buycash
				if (((frmDetail.elements[i+4].value*0) != 0) || ((frmDetail.elements[i+5].value*0) != 0) || ((frmDetail.elements[i+6].value*0) != 0) || ((frmDetail.elements[i+7].value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return false;
				}
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("���û�ǰ�� AGV�������̽��� ���� �Ͻðڽ��ϱ�?") == true) {
			frmDetail.mode.value = "agvregarr";
			frmDetail.action = "/admin/logics/logics_agv_pickup_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("���õ� ��ǰ�� �����ϴ�.");
	}

}

function popViewCurrentStock(itemgubun, itemid, itemoption) {
	var popwin;
	popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popViewCurrentStock','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;

	popwin = window.open('popjumunitem.asp?suplyer=&changesuplyer=Y&shopid=10x10' + '&idx=' + frm.masterid.value,'chulgodetailadd','width=960,height=600,scrollbars=yes,resizable=no');
	popwin.focus();
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,imwdiv){
	if (iidx!='<%= idx %>'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽ� �õ��� �ּ���.');
		return;
	}
<% if Not (C_ADMIN_AUTH) or Not (C_AUTH) then %>
	if (<%= IsExpireEdit %>){
		alert('�δ� ���� ���� ������ ���� �Ұ����մϴ�.');
		return;
	}

	if ((orgexecutedt!='')&&(frmmaster.executedt.value<'<%= BasicMonth %>')){
		alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
		return;
	}
<% end if %>
	var frm;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "itemid") {
			if ((iitemid.indexOf(frm.value + "|") == 0) || (iitemid.indexOf("|" + frm.value + "|") >= 0)) {
				if ((iitemoption.indexOf(frmDetail.elements[i+1].value + "|") == 0) || (iitemoption.indexOf("|" + frmDetail.elements[i+1].value + "|") >= 0)) {
					alert("�ߺ��� ��ǰ�� �ֽ��ϴ�.");
					//popwin.focus();
					return false;
				}
			}
		}
	}

     //��� �⺻ 0��ó��
    var arrsuplycash = isuplycash.split("|");
    isuplycash = "";
    for (i=0;i<arrsuplycash.length;i++){
        if(i==0){
            isuplycash =  parseInt(arrsuplycash[i])*0;
        }else{
        isuplycash = isuplycash + "|" + parseInt(arrsuplycash[i])*0;
        }
    }

	frmDetail.itemgubunarr.value = igubun;
	frmDetail.itemidarr.value = iitemid;
	frmDetail.itemoptionarr.value = iitemoption;
	frmDetail.sellcasharr.value = isellcash;

	frmDetail.suplycasharr.value = isuplycash;
	//frmDetail.suplycasharr.value = isellcash;

	frmDetail.buycasharr.value = ibuycash;
	frmDetail.itemnoarr.value = iitemno;
	frmDetail.itemnamearr.value = iitemname;
	frmDetail.itemoptionnamearr.value = iitemoptionname;
	frmDetail.designerarr.value = iitemdesigner;
	frmDetail.mwdivarr.value = imwdiv;

	frmDetail.mode.value = "adddetail";
	frmDetail.action = "/admin/newstorage/chulgoedit_process.asp";
	frmDetail.submit();
}


//���ڰ��� ǰ�Ǽ� ���
function jsRegEapp(scmidx){
 var frm = document.frmmaster;


	var winEapp = window.open("","popE","width=1000,height=600,scrollbars=yes,resizable=yes");
	document.frmEapp.iSL.value = scmidx;
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
//	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	var winEapp = window.open("/admin/approval/eapp/modeapp.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

//���º���
function jsChangeState(){
    var obj = document.getElementsByName("rdoSt");
    if(confirm("���¸� �����Ͻðڽ��ϱ�?")){
        for(var i=0;i<obj.length;i++){
            if(obj[i].checked){
                document.frmmaster.statecd.value = obj[i].value;
            }
        }

        document.frmmaster.mode.value= "State";
        document.frmmaster.action = "chulgoedit_process.asp";
        document.frmmaster.submit();
    }
}

function ApplyMargin() {
	var frm = document.frmDetail;
	var chk = 0;
	if (!frm.itemid.length) {
		if (frm.chk.checked) {
			frm.suplycash.value = 1 * frm.sellcash.value * (100 - document.frmmaster.storemarginrate.value) / 100;
			chk++;
		}
	} else {
		for (var i=0;i<frm.itemid.length;i++){
			if (frm.chk[i].checked) {
				frm.suplycash[i].value = 1 * frm.sellcash[i].value * (100 - document.frmmaster.storemarginrate.value) / 100;
				chk++;
			}
		}
	}
	if(!chk) {
		alert("�����׸��� �����ϴ�.");
	}
}

function popXL(idx, storemarginrate) {
	var popwin = window.open("/admin/newstorage/pop_chulgodetail_xl_download.asp?idx=" + idx + "&storemarginrate=" + storemarginrate + "&menupos=<%= menupos %>","popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmDetail;

	if (frm.chk.length) {
		for (var i = 0; i < frm.chk.length; i++) {
			frm.chk[i].checked = bool;
			AnCheckClick(frm.chk[i]);
		}
	} else {
		frm.chk.checked = bool;
		AnCheckClick(frm.chk);
	}
}

</script>

<form name="frmEapp" method="post" action="/admin/newstorage/chulgo_regeapp.asp">
<input type="hidden" name="iSL" value="">
</form>
<Form name="frmmaster" method=post action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masterid" value="<%= oipchul.FOneItem.Fid %>">
<input type="hidden" name="code" value="<%= oipchul.FOneItem.Fcode %>">
<input type="hidden" name="storeid" value="<%= oipchul.FOneItem.Fsocid %>">
<input type="hidden" name="socname" value="<%= oipchul.FOneItem.Fsocname%>">
<input type="hidden" name="vatcode" value="008">
<input type="hidden" name="statecd" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">

	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
        	<font color="red"><strong>��ǰ���</strong></font>

		</td>
	</tr>
	<!-- ��ܹ� �� -->
	<tr align="center" bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" >������ڵ�</td>
		<td width="500"  align="left"><%= oipchul.FOneItem.Fcode %></td>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td  align="left"><%= oipchul.FOneItem.Fchargeid %>&nbsp;(<%= oipchul.FOneItem.Fchargename %>)</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" >���ó</td>
		<td align="left"><%= oipchul.FOneItem.Fsocid %>&nbsp;(<%= oipchul.FOneItem.Fsocname %>)</td>
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td  align="left">
			<%if oipchul.FOneItem.Fstatecd = "0" then %>
			<% Call drawSelectBoxIpChulDivcode("etcchulgo", "divcode", oipchul.FOneItem.Fdivcode) %>
			<%else%>
			<%= oipchul.FOneItem.GetDivCodeName %>
			<%end if%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">����û��</td>
		<td align="left">
			<%if oipchul.FOneItem.Fstatecd = "0" then %>
			<input type="text" class="text" name="scheduledt" value="<%= oipchul.FOneItem.Fscheduledt %>" size="10" maxlength=10 readonly><a href="javascript:calendarOpen(frmmaster.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
			<%else%>
			<%= Left(oipchul.FOneItem.Fscheduledt,10) %>
			<%end if%>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td align="Left">
			<%if oipchul.FOneItem.Fstatecd > "0" then %>
			<input type="text" class="text" name=executedt value="<%= Left(oipchul.FOneItem.Fexecutedt,10) %>" size=10 maxlength=10 readonly >
			<a href="javascript:calendarOpen(frmmaster.executedt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
			<% if C_ADMIN_AUTH and (oipchul.FOneItem.Fexecutedt <> "") then %>
			<input type="button" class="button" value="������ں���" onClick="ChChulgoDate(frmmaster)">
			(�����ں�)
			<% end if %>
			<%else%>
				<% if (C_ADMIN_AUTH or C_AUTH) and (oipchul.FOneItem.Fexecutedt <> "") then %>
				<input type="text" class="text" name=executedt value="<%= Left(oipchul.FOneItem.Fexecutedt,10) %>" size=10 maxlength=10 readonly >
				<a href="javascript:calendarOpen(frmmaster.executedt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
				<input type="button" class="button" value="������ں���" onClick="ChChulgoDate(frmmaster)">
				(�����ں�)
				<% else %>
				<%= Left(oipchul.FOneItem.Fexecutedt,10) %>
				<input type="hidden"   name="executedt" value="<%= executedt %>">
				<% end if %>
			<%end if%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�ѼҺ��ڰ�</td>
		<td align="left"><%= FormatNumber(oipchul.FOneItem.Ftotalsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td align="left"><%= FormatNumber(oipchul.FOneItem.Ftotalsuplycash,0) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td align="left">
		  <% IF  oipchul.FOneItem.Fstatecd = "7" or oipchul.FOneItem.Fexecutedt <> "" or not isnull(oipchul.FOneItem.Fexecutedt) THEN  %>
		    <font color="gray">�ֹ����ۼ�</font>> <font color="gray">�ֹ�����</font> > <strong>���Ϸ�</strong>
			<% if C_ADMIN_AUTH and (oipchul.FOneItem.Fexecutedt <> "") then %>
			&nbsp;
			<input type="button" class="button" value="������ȯ" onClick="Chulgo2Jupsu(frmmaster)"> (�����ں�)
			<% end if %>
		  <%elseif oipchul.FOneItem.Fstatecd = "1" then%>
		    <font color="gray">�ֹ����ۼ�</font>> <strong>�ֹ�����</strong> > <font color="gray">���Ϸ�</font>
		  <%ELSE%>
		    <strong>�ֹ����ۼ�</strong>> <font color="gray">�ֹ�����</font> > <font color="gray">���Ϸ�</font>
		  <%END IF%>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">ǰ�ǻ���</td>
		<td align="left">
		  <%if oipchul.FOneItem.pcuserdiv = "900_21" and oipchul.FOneItem.FtplGubun =""   then %>
		      <%  if oipchul.FOneItem.Freportidx = "" or   isNUll( oipchul.FOneItem.Freportidx ) then %>
    		    <strong>�ۼ���</strong> >  <font color="gray">ǰ��������</font> > <font color="gray">ǰ�ǿϷ�</font>
    		  <% elseif oipchul.FOneItem.Freportstate = "7" then %>
    			<font color="gray">�ۼ���</font> >  <font color="gray">ǰ��������</font> > <strong>ǰ�ǿϷ�</strong> (ǰ�ǹ�ȣ: <a href="javascript:jsViewEapp('<%=oipchul.FOneItem.Freportidx%>','<%= oipchul.FOneItem.Freportstate %>');"><%=oipchul.FOneItem.Freportidx%></a>)
    		  <% elseif oipchul.FOneItem.Freportstate = "5" then %>
    			ǰ�ǹݷ� (ǰ�ǹ�ȣ: <a href="javascript:jsViewEapp('<%=oipchul.FOneItem.Freportidx%>','<%= oipchul.FOneItem.Freportstate %>');"><%=oipchul.FOneItem.Freportidx%></a>)
    		  <% else %>
    		    <font color="gray">�ۼ���</font> >   <strong>ǰ��������</strong>  > <font color="gray">ǰ�ǿϷ�</font> (ǰ�ǹ�ȣ: <a href="javascript:jsViewEapp('<%=oipchul.FOneItem.Freportidx%>','<%= oipchul.FOneItem.Freportstate %>');"><%=oipchul.FOneItem.Freportidx%></a>)
    		  <% end if %>

		  <%else%>
		   ǰ�ǹ�����
    	 <%end if%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td colspan="3" align="left"><textarea class="textarea" name="comment" cols=80 rows=6><%= (oipchul.FOneItem.Fcomment) %></textarea></td>
	</tr>

	<!-- �ϴܹ� ���� -->
	 <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if ischulgonotdisp then %>
				<font color="red">���ó�� �ؿܳ� ���ŷ� �����Ǿ� �ִ°��, [OFF]����_������>>�ֹ�����(����)���� ��� �ϼž� �մϴ�.</font><Br>
			<% end if %>
		    <!-- ������ -->
			<%if oipchul.FOneItem.Fstatecd = "0" then '�ۼ��� ���� %>
			    <input type="button" class="button" value="�ӽ�����(�ۼ���)" onclick="tempSave()" <% if ischulgonotdisp then %> disabled<% end if %>>
			    <input type="button" class="button" value="����Ȯ��(�ֹ�����)" onclick="SubmitForm()" <% if ischulgonotdisp then %> disabled<% end if %>>
			<%elseif   oipchul.FOneItem.Fstatecd >= "1" and oipchul.FOneItem.Fstatecd <"7" then %>
			     <input type="button" class="button" value="����" onClick="ModiMaster(frmmaster)" <% if ischulgonotdisp then %> disabled<% end if %>>
			     <% if (C_logics_Part or C_MngPart or C_ADMIN_AUTH) then %>
				    <input type="button" class="button" value="���Ȯ��(���Ϸ�)" onClick="ChulgoMaster(frmmaster)" <% if ischulgonotdisp then %> disabled<% end if %>>
				<%end if%>
			<% end if %>
			 <!-- //������ -->
			  <!-- ǰ�ǻ��� -->

			<% if (oipchul.FOneItem.Fstatecd >= "1" or isNull(oipchul.FOneItem.Fstatecd)) and  oipchul.FOneItem.pcuserdiv = "900_21" and oipchul.FOneItem.FtplGubun ="" then '����Ȯ���� ǰ�� ����(���Ϸ� �Ŀ��� ǰ�ǰ���)%>

    				<%if oipchul.FOneItem.Freportidx = "" OR isNUll( oipchul.FOneItem.Freportidx ) then%>
    				<input type="button" class="button"  value="ǰ�Ǽ� �ۼ�" onClick="jsRegEapp('<%=oipchul.FOneItem.Fid%>');" >
    				<% else %>
    				<input type="button" class="button"  value="ǰ�Ǽ� ����" onClick="jsViewEapp('<%=oipchul.FOneItem.Freportidx%>','<%= oipchul.FOneItem.Freportstate %>');" >
    				<% end if%>
			<% end if%>

			<% if oipchul.FOneItem.Fstatecd < "7" or C_ADMIN_AUTH or C_logics_Part then %>
			<input type="button" class="button" value="��ü����" onClick="DelMaster(frmmaster)" <% if ischulgonotdisp then %> disabled<% end if %>>
			<%end if%>
		</td>
	</tr>
	<!-- �ϴܹ� �� -->
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<%if  (oipchul.FOneItem.Freportidx = "" OR isNUll( oipchul.FOneItem.Freportidx )) then%>���������:
		<input type="text" class="text" style="text-align:right;" id="storemarginrate" name="storemarginrate" value="<%= storemarginrate %>" size="2"> %
		<input type="button" class="button" value="���� ����������" onclick="ApplyMargin()">
		<%end if%>
	</td>
	<td align="right">
		<input type="button" onclick="popXL('<%= idx %>','<%= storemarginrate %>')" value="�����ٿ�" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>

<p style="color:blue;">+ ����� �⺻������ 0������ ��ϵ˴ϴ�. ������ ���Ͻø� ��ǰ�߰� �� [�������ϰ�����]��ư�� �̿��ؼ� ����� �������ּ���</p>

<form name="frmDetail" method="post" action="" style="margin:0px;">
<input type="hidden" name="refergubun" value="B">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="masterid" value="<%= oipchul.FOneItem.Fid %>">
<input type="hidden" name="code" value="<%= oipchul.FOneItem.Fcode %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
<input type="hidden" name="alinkcode" value="<%= oipchul.FOneItem.Falinkcode %>">
<input type="hidden" name="currencyUnit" value="<%= currencyUnit %>">
<input type="hidden" name="currencyUnit_Pos" value="<%= currencyUnit_Pos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="14">
			<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
			        	<font color="red"><strong>�󼼳���</strong></font>
			        	&nbsp;&nbsp;
			        	<font color="#EE4444">����</font>&nbsp;��Ź&nbsp;<font color="#4444EE">��ü���</font>
	        		</td>
	        		<td align="right">
	        			�ѰǼ�:  <%= oipchuldetail.FResultCount %>
			        	&nbsp;
			        	<%if (oipchul.FOneItem.Freportidx = "" OR isNUll( oipchul.FOneItem.Freportidx )) and  oipchul.FOneItem.Fstatecd < "7" then%>
			        	<input type="button" class="button" value=" ��ǰ�߰� " onClick="AddItems(frmmaster)" <% if ischulgonotdisp then %> disabled<% end if %>>
			        	<%end if%>
	        		</td>
	        	</tr>
	        </table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=20><Input Type="checkbox" name="ckall" onClick="ckAll(this)"></td>
		<td width="120">��ǰ�ڵ�</td>
		<td width=80>�귣��ID</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width=50>����</td>
		<td width=70>�ǸŰ�</td>
		<td width=70>���</td>
		<td width=70>���԰�</td>
		<td width=40>���<br>������</td>
		<td width=40>����<br>����</td>
		<td width=25>���<br>����</td>
		<td width=25>����<br>����</td>
		<td width=25>����<br>����<br>����</td>
	</tr>
<% for i=0 to oipchuldetail.FResultCount -1 %>
	<%
	ttlitemno = ttlitemno + oipchuldetail.FItemList(i).Fitemno
	sellcashtotal = sellcashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsellcash
	suplycashtotal = suplycashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsuplycash
	buycashtotal = buycashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fbuycash
	%>
	<tr bgcolor="#FFFFFF">
		<td>
			<input type=checkbox name=chk value="<%= i %>" onClick="AnCheckClick(this);">
			<input type="hidden" name="itemgubun" value="<%= oipchuldetail.FItemList(i).fitemgubun %>">
			<input type="hidden" name="itemid" value="<%= oipchuldetail.FItemList(i).FItemId %>">
			<input type="hidden" name="itemoption" value="<%= oipchuldetail.FItemList(i).FItemOption %>">
		</td>
		<td>
			<a href="javascript:popViewCurrentStock('<%= oipchuldetail.FItemList(i).Fiitemgubun %>', '<%= oipchuldetail.FItemList(i).FItemId %>', '<%= oipchuldetail.FItemList(i).FItemOption %>');">
				<%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %>-<%= oipchuldetail.FItemList(i).FItemOption %>
			</a>
		</td>
		<td><%= oipchuldetail.FItemList(i).Fimakerid %></td>
		<td><%= oipchuldetail.FItemList(i).FIItemName %></td>
		<td align=center><%= oipchuldetail.FItemList(i).FIItemoptionName %></td>
		<td align=center><input type="text" class="text" name="itemno" value="<%= oipchuldetail.FItemList(i).Fitemno %>" size=4 maxlength=6></td>
		<td align=right><input type="text" class="text" name="sellcash" value="<%= oipchuldetail.FItemList(i).Fsellcash %>" size=7 maxlength=9 style="text-align:right"></td>
		<td align=right><input type="text" class="text" name="suplycash" value="<%= oipchuldetail.FItemList(i).Fsuplycash %>" size=7 maxlength=9 style="text-align:right"></td>
		<td align=right><input type="text" class="text" name="buycash" value="<%= oipchuldetail.FItemList(i).Fbuycash %>" size=7 maxlength=9 style="text-align:right"></td>
		<td align=center>
		<% if oipchuldetail.FItemList(i).Fsellcash<>0 then %>
		<%= 100-CLng(oipchuldetail.FItemList(i).Fsuplycash/oipchuldetail.FItemList(i).Fsellcash*100*100)/100 %>%
		<% end if %>
		</td>
		<td align=center>
		<% if oipchuldetail.FItemList(i).Fsellcash<>0 then %>
		<%= 100-CLng(oipchuldetail.FItemList(i).Fbuycash/oipchuldetail.FItemList(i).Fsellcash*100*100)/100 %>%
		<% end if %>
		</td>
		<td align="center"><%= oipchuldetail.FItemList(i).FMWgubun %></td>
		<% if (C_ADMIN_AUTH) and ((oipchuldetail.FItemList(i).FOnlineMwdiv="W") and (oipchuldetail.FItemList(i).FMWgubun<>"C")) or (oipchuldetail.FItemList(i).FOnlineMwdiv="U") then %>
		<td align="center"><a href="javascript:ConvertWiChulgo('<%= oipchuldetail.FItemList(i).Fid %>');"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FOnlineMwdiv %></font></a></td>
		<td align="center"><a href="javascript:ConvertWiChulgo('<%= oipchuldetail.FItemList(i).Fid %>');"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FCenterMwdiv %></font></a></td>
		<% else %>
		<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FOnlineMwdiv %></font></td>
		<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FCenterMwdiv %></font></td>
		<% end if %>
		<input type="hidden" name="didx" value="<%= oipchuldetail.FItemList(i).Fid %>">
	</tr>
	<% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=5>�Ѱ�</td>
		<td align=center><%= FormatNumber(ttlitemno,0) %></td>
		<td align=right><b><%= FormatNumber(sellcashtotal,0) %></b></td>
		<td align=right><b><%= FormatNumber(suplycashtotal,0) %></b></td>
		<td align=right><b><%= FormatNumber(buycashtotal,0) %></b></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%if (C_ADMIN_AUTH) or (C_AUTH) or (C_MngPart) or (C_OP_AUTH) or (oipchul.FOneItem.Freportidx = "" OR isNUll( oipchul.FOneItem.Freportidx )) and  oipchul.FOneItem.Fstatecd < "7" or (idx = 355771)   then%>
				<input type="button" class="button" value=" ���û�ǰ���� " onclick="ModiDetail(frmmaster,frmDetail)" <% if ischulgonotdisp then %> disabled<% end if %>>
				<input type="button" class="button" value=" ���û�ǰ���� " onclick="DelDetail(frmmaster,frmDetail)" <% if ischulgonotdisp then %> disabled<% end if %>>
			<%end if%>
			<% if oipchul.FOneItem.Falinkcode="" or isnull(oipchul.FOneItem.Falinkcode) then %>
				<input type="button" class="button" value=" ���û�ǰAGV�������̽�����" onclick="regAGVArr();">
			<% end if %>
		</td>
	</tr>
</table>
</form>

<% if ochulgolog.FResultCount >0 then %>
<br><br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF" height="25">
	<td colspan="19">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<font color="red"><strong>���� �α�</strong></font>
		&nbsp;&nbsp;&nbsp;�ִ� 50������ ���� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">����</td>
	<td>�������</td>
	<td width="60">�ֹ��ڵ�</td>
	<td width="100">���óID</td>
	<td width="100">���ó��</td>
	<td width="60">�����</td>
	<td width="60">ó����</td>
	<td width="60">������</td>
	<td width="60">ǰ�ǻ���</td>
	<td width="70">��û��</td>
	<td width="70">�����</td>
	<td width="70">�ǸŰ�</td>
	<td width="70">���</td>
	<td width="70">���԰�</td>
	<td width="60">����</td>
	<td width="80">����</td>
	<td width="40">������</td>
	<td width="40">����</td>
</tr>

<% if ochulgolog.FResultCount >0 then %>
	<% for i=0 to ochulgolog.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=24>
		<td align="center">
			<%= left(ochulgolog.FItemList(i).flogregdate,10) %>
			<br><%= mid(ochulgolog.FItemList(i).flogregdate,12,22) %>
			<br><%= ochulgolog.FItemList(i).flogadminid %>
		</td>
		<td align="left"><%= ochulgolog.FItemList(i).fbigo %></td>
		<td align=center><%= ochulgolog.FItemList(i).Falinkcode %></td>
		<td align=left><%= ochulgolog.FItemList(i).Fsocid %></td>
		<td align=left><%= ochulgolog.FItemList(i).Fsocname %></b></td>
		<td align=center><%= ochulgolog.FItemList(i).Fchargename %></td>
		<td align=center><%= ochulgolog.FItemList(i).Ffinishname %></td>
		<td align=center>
		    <% IF ochulgolog.FItemList(i).Fstatecd = "7" or ochulgolog.FItemList(i).Fexecutedt <> "" or not isnull(ochulgolog.FItemList(i).Fexecutedt) THEN  %>
		    	���Ϸ�
		    <%elseif ochulgolog.FItemList(i).Fstatecd = "1" then%>
		    	�ֹ�����
		    <%ELSE%>
		    	�ֹ����ۼ�
		    <%END IF%>
		</td>
		<td align=center>
			<%if ochulgolog.FItemList(i).Freportidx <> "" and not isNUll( ochulgolog.FItemList(i).Freportidx ) then%>
				<%if ochulgolog.FItemList(i).Freportstate = "7" then %>
					ǰ�ǿϷ�
				<%elseif ochulgolog.FItemList(i).Freportstate = "5" then %>
					ǰ�ǹݷ�
				<%else%>
					������
				<%end if%>
			<% end if%>
		</td>
		<td align=center><%= Left(ochulgolog.FItemList(i).Fscheduledt,10) %></td>
		<td align=center><%= Left(ochulgolog.FItemList(i).Fexecutedt,10) %></td>
		<td align=right><%= FormatNumber(ochulgolog.FItemList(i).Ftotalsellcash,0) %></td>
		<td align="right"><%= FormatNumber(ochulgolog.FItemList(i).Ftotalsuplycash,0) %></td>
		<td align="right"><%= FormatNumber(ochulgolog.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=right><%= FormatNumber(ochulgolog.FItemList(i).ftotalitemno,0) %></td>
		<td align=center><%= ochulgolog.FItemList(i).GetDivCodeName %></td>
		<td align=right>
			<% if ochulgolog.FItemList(i).Ftotalsellcash<>0 then %>
				<%= 100-CLng(ochulgolog.FItemList(i).Ftotalsuplycash/ochulgolog.FItemList(i).Ftotalsellcash*100*100)/100 %>%
			<% end if %>
		</td>
		<td align=right>
			<% if ochulgolog.FItemList(i).Ftotalsuplycash<>0 then %>
				<%= round((100-CLng(ochulgolog.FItemList(i).Ftotalbuycash/ochulgolog.FItemList(i).Ftotalsuplycash*100*100)/100),2) %>%
			<% end if %>
		</td>
	</tr>
	<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>
<% end if %>

<%
set oipchuldetail = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
