<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ���ʽ� ����
' History : 2011.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/bonuscoupon/bonuscoupon_cls.asp" -->

<%
dim ocoupon , shopidchoice ,idx ,coupontype ,couponvalue ,couponname ,startdate ,expiredate
dim isusing,minbuyprice ,targetitemlist, targetbrandlist, openfinishdate ,etcstr ,isopenlistcoupon, couponmeaipprice
dim validsitename ,doublesaleyn ,limityn , limitno ,openfinishdateTime ,startdatetime ,expiredateTime ,oshop ,lastupdateadminid ,i
dim exitemidlist, exbrandidlist
dim arrexitemidlist, arrexbrandidlist
dim IsTargetItemCoupon, IsTargetBrandCoupon
dim usecondition
	idx = requestCheckVar(request("idx"),10)

if idx="" then idx=0

'/������
set ocoupon = new CCouponlist
	ocoupon.FRectIdx = idx

	'/������ ��� ���� ������
	if idx<>0 then
		ocoupon.GetCouponMasteritem

		if ocoupon.ftotalcount > 0 then
			idx = ocoupon.FOneItem.Fidx
			coupontype = ocoupon.FOneItem.Fcoupontype
			couponvalue = ocoupon.FOneItem.Fcouponvalue
			couponname = ocoupon.FOneItem.Fcouponname
			startdate = ocoupon.FOneItem.Fstartdate
			expiredate = ocoupon.FOneItem.Fexpiredate
			isusing = ocoupon.FOneItem.Fisusing
			minbuyprice = ocoupon.FOneItem.Fminbuyprice

			targetitemlist = ocoupon.FOneItem.Ftargetitemlist
			targetbrandlist = ocoupon.FOneItem.Ftargetbrandlist

			openfinishdate = ocoupon.FOneItem.FOpenFinishDate
			etcstr = ocoupon.FOneItem.Fetcstr
			isopenlistcoupon = ocoupon.FOneItem.Fisopenlistcoupon
			couponmeaipprice = ocoupon.FOneItem.Fcouponmeaipprice
			validsitename = ocoupon.FOneItem.Fvalidsitename
			doublesaleyn = ocoupon.FOneItem.fdoublesaleyn
			limityn = ocoupon.FOneItem.flimityn
			limitno = ocoupon.FOneItem.flimitno
			lastupdateadminid = ocoupon.FOneItem.flastupdateadminid

			exitemidlist = ocoupon.FOneItem.Fexitemidlist
			exbrandidlist = ocoupon.FOneItem.Fexbrandidlist

			if IsNull(exitemidlist) then
				exitemidlist = ""
			end if

			if IsNull(exbrandidlist) then
				exbrandidlist = ""
			end if

			IsTargetItemCoupon = ocoupon.FOneItem.IsTargetItemCoupon
			IsTargetBrandCoupon = ocoupon.FOneItem.IsTargetBrandCoupon

			usecondition = ""
			if (IsTargetItemCoupon) then
				usecondition = "I"
			end if
			if (IsTargetBrandCoupon) then
				usecondition = "B"
			end if

			startdatetime = Num2Str(Hour(startdate),2,"0","R") & ":" & Num2Str(Minute(startdate),2,"0","R")& ":" & Num2Str(second(startdate),2,"0","R")
			expiredateTime = Num2Str(Hour(expiredate),2,"0","R") & ":" & Num2Str(Minute(expiredate),2,"0","R")& ":" & Num2Str(second(expiredate),2,"0","R")
			openfinishdateTime = Num2Str(Hour(openfinishdate),2,"0","R") & ":" & Num2Str(Minute(openfinishdate),2,"0","R")& ":" & Num2Str(second(openfinishdate),2,"0","R")
		end if
	end if

	arrexitemidlist = Split(exitemidlist, ",")
	arrexbrandidlist = Split(exbrandidlist, ",")

'/��������
set oshop = new CCouponlist
	oshop.FRectIdx = idx

	if idx<>0 then
		oshop.GetCouponshopList
	end if

if startdate="" then startdate=date
if startdatetime="" then startdatetime="00:00:00"
if expiredate="" then expiredate=dateAdd("d",1,date)
if expiredateTime="" then expiredateTime="23:59:59"
if openfinishdate="" then openfinishdate=dateAdd("d",1,date)
if openfinishdateTime="" then openfinishdateTime="23:59:59"
if doublesaleyn = "" then doublesaleyn = "N"
if limityn = "" then limityn = "Y"
if validsitename = "" then validsitename = "10X10OFFLINE"
%>

<script type='text/javascript'>

function CheckVallidNumber(obj, objname) {
	if (obj.value.length < 1) {
		alert(objname + '�� �Է��ϼ���.');
		obj.focus();
		return false;
	}

	if (obj.value*0 != 0){
		alert(objname + '�� ���ڸ� �Է��ϼ���.');
		obj.focus();
		return false;
	}

	if (obj.value*1 < 0){
		alert(objname + '�� 0 ���� ���� �� �����ϴ�.');
		obj.focus();
		return false;
	}

	return true;
}

function submitForm(frm){
	if (frm.couponname.value.length<1){
		alert('�������� �Է��ϼ���.');
		frm.couponname.focus();
		return;
	}

    if ((!frm.coupontype[0].checked)&&(!frm.coupontype[1].checked)){
        alert('���� Ÿ���� �����ϼ���.');
		frm.coupontype[0].focus();
		return;
    }

	if (CheckVallidNumber(frm.minbuyprice, "�ּ� ���űݾ�") != true) {
		return;
	}

	if (CheckVallidNumber(frm.couponvalue, "���� �ݾ�") != true) {
		return;
	}

	if (frm.startdate.value.length<1){
		alert('��ȿ�Ⱓ �������� �Է��ϼ���.');
		frm.startdate.focus();
		return;
	}

	if (frm.expiredate.value.length<1){
		alert('��ȿ�Ⱓ �������� �Է��ϼ���.');
		frm.expiredate.focus();
		return;
	}

	if (frm.openfinishdate.value.length<1){
		alert('���� �߱� �������� �Է��ϼ���.');
		frm.openfinishdate.focus();
		return;
	}

	if (frm.shopid == undefined) {
		alert('��������� �Է��ϼ���.');
		frm.shopidchoice.focus();
		return;
	}

	if ((frm.coupontype[0].checked == true) && (frm.couponvalue.value*1 > 15)) {
		// ������
		alert('15% �� �Ѵ� ���������� ������ �� �����ϴ�.');
		frm.couponvalue.focus();
		return;
	}

	if ((frm.coupontype[1].checked == true) && (frm.couponvalue.value*1 > frm.minbuyprice.value*0.2)) {
		// ������
		alert('�������ξ��� �ּұ��űݾ��� 20% �� ���� �� �����ϴ�.');
		frm.couponvalue.focus();
		return;
	}

	/*
	if ((frm.coupontype[1].checked == true) && (frm.usecondition.value != "I")) {
		// ȯ�ҽ� ������ �ȴ�.
		alert('���������� ��Ź��ǰ�� ���ؼ��� ������ �� �ֽ��ϴ�.');
		return;
	}
	*/

	if (frm.usecondition.value == "I") {
		if (frm.targetitemlist.value == "") {
			alert('�����ǰ�� �����ϼ���.');
			return;
		}

		var shopidcount = 0;
		if (frm.shopid != undefined) {
			shopidcount = 1;

			if (frm.shopid.length != undefined) {
				shopidcount = frm.shopid.length;
			}
		}

		if (shopidcount != 1) {
			if (shopidcount < 1) {
				alert("���� �����ϼ���");
			} else {
				alert("��ǰ�� ������ ������ ���� �ϳ��� ���� ������ �� �ֽ��ϴ�.");
			}
			return;
		}
	}

	var exitemidcount = 0;
	if (frm.exitemid != undefined) {
		exitemidcount = 1;

		if (frm.exitemid.length != undefined) {
			exitemidcount = frm.exitemid.length;
		}
	}

	if (exitemidcount > 10) {
		alert("10���� �ʰ��Ͽ� ���ܻ�ǰ�� ������ �� �����ϴ�.");
		return;
	}

	var exbrandidcount = 0;
	if (frm.exbrandid != undefined) {
		exbrandidcount = 1;

		if (frm.exbrandid.length != undefined) {
			exbrandidcount = frm.exbrandid.length;
		}
	}

	if (exbrandidcount > 10) {
		alert("10���� �ʰ��Ͽ� ���ܺ귣�带 ������ �� �����ϴ�.");
		return;
	}

	if ((frm.usecondition.value == "B") && (frm.targetbrandlist.value == "")) {
		alert('����귣�带 �����ϼ���.');
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		if (frm.usecondition.value == "I") {
			frm.targetbrandlist.value = "";
		}

		if (frm.usecondition.value == "B") {
			frm.targetitemlist.value = "";
		}

		frm.submit();
	}
}

function EnableBox(comp){
	if (comp.checked){
		frm.targetitemlist.disabled = false;
		frm.couponmeaipprice.disabled = false;

		frm.targetitemlist.style.backgroundColor = "#FFFFFF";
		frm.couponmeaipprice.style.backgroundColor = "#FFFFFF";
	}else{
		frm.targetitemlist.disabled = true;
		frm.couponmeaipprice.disabled = true;

		frm.targetitemlist.style.backgroundColor = "#E6E6E6";
		frm.couponmeaipprice.style.backgroundColor = "#E6E6E6";
	}

}

function SetLimitNo(v) {
	if (v == 'S') {
		frm.limitno.readonly = false;
		frm.limitno.style.backgroundColor = "#FFFFFF";
	} else if (v == 'N') {
		frm.limitno.readonly = true;
		frm.limitno.style.backgroundColor = "#E6E6E6";
		frm.limitno.value = '0';
	} else {
		frm.limitno.style.backgroundColor = "#E6E6E6";
		frm.limitno.readonly = true;
		frm.limitno.value = '1';
	}
}

//tr�߰�
function AutoInsert() {

	if (frm.shopidchoice.value==""){
		alert('������ ������ �ּ���');
		frm.shopidchoice.focus();
		return;
	}
	var choice = frm.shopidchoice.value;
	var f = document.all;

	var rowLen = f.div1.rows.length;
	var r  = f.div1.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<input type='hidden' name='shopid' value='"+choice+"'> &nbsp; "+choice+" &nbsp; <img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>";
	c0.innerHTML = inHtml;
	frm.tmpshopid.value = parseInt(frm.tmpshopid.value) + 1
}

function InsertExItemID() {
	if (frm.exitemidchoice.value==""){
		alert('���� ��ǰ�ڵ带 �˻��ϼ���.');
		return;
	}

	var choice = frm.exitemidchoice.value;
	var f = document.all;

	var rowLen = f.divexitemidlist.rows.length;
	var r  = f.divexitemidlist.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<input type='hidden' name='exitemid' value='"+choice+"'> &nbsp; "+choice+" &nbsp; <img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearExItemRow(this)'>";
	c0.innerHTML = inHtml;
}

function InsertExBrandID() {
	var choice = frm.exbrandidchoice.value;
	var f = document.all;

	var rowLen = f.divexbrandidlist.rows.length;
	var r  = f.divexbrandidlist.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<input type='hidden' name='exbrandid' value='"+choice+"'> &nbsp; "+choice+" &nbsp; <img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearExBrandRow(this)'>";
	c0.innerHTML = inHtml;
}

//tr����
function clearRow(tdObj) {
	if(confirm("�����Ͻ� ���� �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);

		document.frm.targetitemlist.value = "";
	} else {
		return false;
	}
}

function clearExItemRow(tdObj) {
	if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

function clearExBrandRow(tdObj) {
	if(confirm("�����Ͻ� �귣�带 �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

function UseConditionChanged(frm) {
	var trtargetitemlist = document.getElementById("trtargetitemlist");
	var trtargetbrandlist = document.getElementById("trtargetbrandlist");

	trtargetitemlist.style.display = 'none';
	trtargetbrandlist.style.display = 'none';

	if (frm.usecondition.value == "I") {
		trtargetitemlist.style.display = 'block';
	}

	if (frm.usecondition.value == "B") {
		trtargetbrandlist.style.display = 'block';
	}
}

function jsSearchItemID(frm, frmname, targetinputboxname) {
	var shopidcount = 0;
	var shopid = "";

	if (frm.shopid != undefined) {
		shopidcount = 1;

		if (frm.shopid.length != undefined) {
			shopidcount = frm.shopid.length;
			shopid = frm.shopid[0].value;
		} else {
			shopid = frm.shopid.value;
		}
	}

	if (shopidcount != 1) {
		if (shopidcount < 1) {
			alert("���� �����ϼ���");
		} else {
			alert("��ǰ�� ������ �����Ϸ��� �� ���庰�� ����ؾ� �մϴ�.");
		}
		return;
	}

	var popwin;
	popwin = window.open("/common/offshop/pop_itemSelectOne_off.asp?shopid=" + shopid + "&frmname=" + frmname + "&targetinputboxname=" + targetinputboxname, "jsSearchItemID", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>

<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>

<style>
	.display_date {cursor:pointer; display:inline-block; font-family: "Verdana", "����"; font-size: 9pt; background-color: #FFFFFF; border:1px solid #BABABA; color: #000000; width:85px; height: 20px; padding:0 0 1px 2px;}
</style>

<table width="900" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/offshop/bonuscoupon/coupon_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="idx" value="<%=idx%>">
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="120">IDX</td>
	<td bgcolor="#FFFFFF"><%= idx %></td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="validsitename" value="10X10OFFLINE" <%= CHKIIF(validsitename="10X10OFFLINE","checked","") %>>�ٹ����� ����
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="couponname" value="<%= couponname %>" maxlength="100" size=40>
		&nbsp;
		(ex �ٹ����� �ָ� ����)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">����Ÿ��</td>
	<td bgcolor="#FFFFFF">
		�ߺ����� : <% DrawDoubleSaleYN "doublesaleyn" ,doublesaleyn, "", "" %>
		&nbsp;
		��ü�߱���� : <% DrawLimitYN "limityn",limityn," onchange='SetLimitNo(this.value);'","" %>
		<input type="text" name="limitno" size=6 maxlength=10 value="<%= limitno %>">
		<script language="javascript">
			SetLimitNo('<%= limityn %>')
		</script>
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td bgcolor="#FFFFFF"><% DrawUseCondition "usecondition" , usecondition, " onChange='UseConditionChanged(frm)' ", "Y" %> &nbsp; <input type="text" name=minbuyprice value="<%= minbuyprice %>" maxlength=7 size=10  >�� �̻� ���Ž�(����)</td>
</tr>
<tr height=30 id="trtargetitemlist" style="display:<% if IsTargetItemCoupon then %>block<% else %>none<% end if %>">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">�����ǰ����</td>
	<td bgcolor="#FFFFFF">
		��ǰ�ڵ�: <input type=text name=targetitemlist value="<%= targetitemlist %>" size=14 maxlength=14 readonly style='background-color:#E6E6E6;'>
		<input type="button" onClick="jsSearchItemID(frm, this.form.name,'targetitemlist')" value="�˻�" class='button'>
		(��Ź ��ǰ�� ���ε�, ��Ź ��ǰ�� �ɼ� ���ο� �����)
	</td>
</tr>
<tr height=30 id="trtargetbrandlist" style="display:<% if IsTargetBrandCoupon then %>block<% else %>none<% end if %>">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">����귣������</td>
	<td bgcolor="#FFFFFF">
		�� �� ��: <input type=text name=targetbrandlist value="<%= targetbrandlist %>" size=32 maxlength=32 readonly style='background-color:#E6E6E6;'>
		<input type="button" class="button" value="�귣��˻�" onclick="jsSearchBrandID(this.form.name,'targetbrandlist');" >
		(��Ź �귣�常 ���ε�)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">��Ź��ǰ <font color=red>����</font></td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="divexitemidlist" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				��ǰ�ڵ�: <input type=text name=exitemidchoice value="" size=14 maxlength=14 readonly style='background-color:#E6E6E6;'>
				<input type="button" onClick="jsSearchItemID(frm, this.form.name,'exitemidchoice')" value="�˻�" class='button'>
				<input type="button" onClick="InsertExItemID()" value="�߰�" class='button'>
				(��ǰ�� �ɼ� ���ο� �����)
			</td>
		</tr>
		<% if exitemidlist <> "" then %>
		<% for i = 0 to Ubound(arrexitemidlist) %>
		<tr>
			<td>
				<input type="hidden" name="exitemid" value="<%= arrexitemidlist(i) %>">
				&nbsp; <%= arrexitemidlist(i) %> &nbsp;
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
			</td>
		</tr>
		<% next %>
		<% end if %>
		</table>
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">��Ź�귣�� <font color=red>����</font></td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="divexbrandidlist" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				�귣��: <input type=text name=exbrandidchoice value="" size=20 maxlength=32 readonly style='background-color:#E6E6E6;'>
				<input type="button" onClick="jsSearchBrandID(this.form.name,'exbrandidchoice')" value="�˻�" class='button'>
				<input type="button" onClick="InsertExBrandID()" value="�߰�" class='button'>
			</td>
		</tr>
		<% if exbrandidlist <> "" then %>
		<% for i = 0 to Ubound(arrexbrandidlist) %>
		<tr>
			<td>
				<input type="hidden" name="exbrandid" value="<%= arrexbrandidlist(i) %>">
				&nbsp; <%= arrexbrandidlist(i) %> &nbsp;
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
			</td>
		</tr>
		<% next %>
		<% end if %>
		</table>
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isopenlistcoupon" value="N" <% if isopenlistcoupon="N" or isopenlistcoupon="" then Response.Write "checked" %>>��ü��
		<input type="radio" name="isopenlistcoupon" value="Y" <% if isopenlistcoupon="Y" then Response.Write "checked" %>>���ð�(������)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">�����������</td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="div1" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				<!-- ����,����,�ؿ��� = 1,3,7 -->
				<% drawSelectBoxOffShopdiv_off "shopidchoice",shopidchoice , "1" ,"","" %>
				<input type="button" onClick="AutoInsert()" value="�߰�" class='button'>
				<input type="hidden" name="tmpshopid" value=0>
			</td>
		</tr>
		<% if oshop.fresultcount > 0 then %>
		<% for i = 0 to oshop.fresultcount -1 %>
		<tr>
			<td>
				<input type="hidden" name="shopid" value="<%= oshop.fitemlist(i).fshopid %>">
				&nbsp; <%= oshop.fitemlist(i).fshopid %> &nbsp;
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
			</td>
		</tr>
		<% next %>
		<% end if %>
		</table>
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="couponvalue" value="<%= couponvalue %>" maxlength=7 size=10>
    	<% if coupontype="1" then %>
    		<input type="radio" name="coupontype" value="1" checked >%����
    		<input type="radio" name="coupontype" value="2" >������
    	<% elseif coupontype="2" then %>
    		<input type="radio" name="coupontype" value="1" >%����
    		<input type="radio" name="coupontype" value="2" checked >������
    	<% else %>
    		<input type="radio" name="coupontype" value="1" >%����
    		<input type="radio" name="coupontype" value="2" checked >������
    	<% end if %>
		(�ݾ� �Ǵ� % ����)
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">��ȿ�Ⱓ</td>
	<td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="startdate" value="<%=left(startdate,10)%>" size=10 readonly ><a href="javascript:calendarOpen(frm.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    	<input type="text" name="startdateTime" size="8" maxlength="8" class="text" value="<%=startdateTime%>">
    	~
    	<input type="text" class="text" name="expiredate" value="<%=left(expiredate,10)%>" size=10 readonly ><a href="javascript:calendarOpen(frm.expiredate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    	<input type="text" name="expiredateTime" size="8" maxlength="8" class="text" value="<%=expiredateTime%>">
	    (<%= Left(now(),10) %> 00:00:00 ~ <%= Left(now(),10) %> 23:59:59)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">�����߱޸�����</td>
	<td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="openfinishdate" value="<%=left(openfinishdate,10)%>" size=10 readonly ><a href="javascript:calendarOpen(frm.openfinishdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    	<input type="text" name="openfinishdateTime" size="8" maxlength="8" class="text" value="<%=openfinishdateTime%>">
		(<%= Left(now(),10) %> 23:59:59)
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ�ڸ�Ʈ</td>
	<td bgcolor="#FFFFFF"><textarea name="etcstr" cols=80 rows=8><%= etcstr %></textarea></td>
</tr>

<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" maxlength=7 size=10 <% if IsUsing="Y" or IsUsing="" then response.write " checked" %>>Y
		<input type="radio" name="isusing" value="N" maxlength=7 size=10>N
	</td>
</tr>
<tr height=30>
	<td colspan="2" align=center bgcolor="#FFFFFF">
		<input type=button value="����" onClick="submitForm(frm);" class="button">
		<input type=button value="�������" onClick="location.href='/admin/offshop/bonuscoupon/couponlist.asp?menupos=<%=menupos%>';" class="button">
	</td>
</tr>
</form>
</table>



<%
set ocoupon = Nothing
set oshop = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->