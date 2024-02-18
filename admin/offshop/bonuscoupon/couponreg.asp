<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  보너스 쿠폰
' History : 2011.05.12 한용민 생성
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

'/쿠폰상세
set ocoupon = new CCouponlist
	ocoupon.FRectIdx = idx

	'/수정일 경우 내역 가져옴
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

'/매장정보
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
		alert(objname + '을 입력하세요.');
		obj.focus();
		return false;
	}

	if (obj.value*0 != 0){
		alert(objname + '에 숫자를 입력하세요.');
		obj.focus();
		return false;
	}

	if (obj.value*1 < 0){
		alert(objname + '은 0 보다 작을 수 없습니다.');
		obj.focus();
		return false;
	}

	return true;
}

function submitForm(frm){
	if (frm.couponname.value.length<1){
		alert('쿠폰명을 입력하세요.');
		frm.couponname.focus();
		return;
	}

    if ((!frm.coupontype[0].checked)&&(!frm.coupontype[1].checked)){
        alert('쿠폰 타입을 선택하세요.');
		frm.coupontype[0].focus();
		return;
    }

	if (CheckVallidNumber(frm.minbuyprice, "최소 구매금액") != true) {
		return;
	}

	if (CheckVallidNumber(frm.couponvalue, "할인 금액") != true) {
		return;
	}

	if (frm.startdate.value.length<1){
		alert('유효기간 시작일을 입력하세요.');
		frm.startdate.focus();
		return;
	}

	if (frm.expiredate.value.length<1){
		alert('유효기간 만료일을 입력하세요.');
		frm.expiredate.focus();
		return;
	}

	if (frm.openfinishdate.value.length<1){
		alert('쿠폰 발급 마감일을 입력하세요.');
		frm.openfinishdate.focus();
		return;
	}

	if (frm.shopid == undefined) {
		alert('적용매장을 입력하세요.');
		frm.shopidchoice.focus();
		return;
	}

	if ((frm.coupontype[0].checked == true) && (frm.couponvalue.value*1 > 15)) {
		// 사고방지
		alert('15% 를 넘는 할인쿠폰은 생성할 수 없습니다.');
		frm.couponvalue.focus();
		return;
	}

	if ((frm.coupontype[1].checked == true) && (frm.couponvalue.value*1 > frm.minbuyprice.value*0.2)) {
		// 사고방지
		alert('정액할인액이 최소구매금액의 20% 를 넘을 수 없습니다.');
		frm.couponvalue.focus();
		return;
	}

	/*
	if ((frm.coupontype[1].checked == true) && (frm.usecondition.value != "I")) {
		// 환불시 문제가 된다.
		alert('정액할인은 위탁상품에 대해서만 적용할 수 있습니다.');
		return;
	}
	*/

	if (frm.usecondition.value == "I") {
		if (frm.targetitemlist.value == "") {
			alert('적용상품을 지정하세요.');
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
				alert("샵을 지정하세요");
			} else {
				alert("상품에 쿠폰을 적용할 경우는 하나의 샵만 지정할 수 있습니다.");
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
		alert("10개를 초과하여 제외상품을 선택할 수 없습니다.");
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
		alert("10개를 초과하여 제외브랜드를 선택할 수 없습니다.");
		return;
	}

	if ((frm.usecondition.value == "B") && (frm.targetbrandlist.value == "")) {
		alert('적용브랜드를 지정하세요.');
		return;
	}

	var ret = confirm('저장 하시겠습니까?');

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

//tr추가
function AutoInsert() {

	if (frm.shopidchoice.value==""){
		alert('매장을 선택해 주세요');
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
		alert('먼저 상품코드를 검색하세요.');
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

//tr삭제
function clearRow(tdObj) {
	if(confirm("선택하신 샵을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);

		document.frm.targetitemlist.value = "";
	} else {
		return false;
	}
}

function clearExItemRow(tdObj) {
	if(confirm("선택하신 상품을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

function clearExBrandRow(tdObj) {
	if(confirm("선택하신 브랜드를 삭제하시겠습니까?") == true) {
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
			alert("샵을 지정하세요");
		} else {
			alert("상품에 쿠폰을 적용하려면 각 매장별로 등록해야 합니다.");
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
	.display_date {cursor:pointer; display:inline-block; font-family: "Verdana", "돋움"; font-size: 9pt; background-color: #FFFFFF; border:1px solid #BABABA; color: #000000; width:85px; height: 20px; padding:0 0 1px 2px;}
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
	<td bgcolor="<%= adminColor("tabletop") %>">쿠폰구분</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="validsitename" value="10X10OFFLINE" <%= CHKIIF(validsitename="10X10OFFLINE","checked","") %>>텐바이텐 매장
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">쿠폰명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="couponname" value="<%= couponname %>" maxlength="100" size=40>
		&nbsp;
		(ex 텐바이텐 주말 쿠폰)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">쿠폰타입</td>
	<td bgcolor="#FFFFFF">
		중복할인 : <% DrawDoubleSaleYN "doublesaleyn" ,doublesaleyn, "", "" %>
		&nbsp;
		전체발급장수 : <% DrawLimitYN "limityn",limityn," onchange='SetLimitNo(this.value);'","" %>
		<input type="text" name="limitno" size=6 maxlength=10 value="<%= limitno %>">
		<script language="javascript">
			SetLimitNo('<%= limityn %>')
		</script>
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">사용조건</td>
	<td bgcolor="#FFFFFF"><% DrawUseCondition "usecondition" , usecondition, " onChange='UseConditionChanged(frm)' ", "Y" %> &nbsp; <input type="text" name=minbuyprice value="<%= minbuyprice %>" maxlength=7 size=10  >원 이상 구매시(숫자)</td>
</tr>
<tr height=30 id="trtargetitemlist" style="display:<% if IsTargetItemCoupon then %>block<% else %>none<% end if %>">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">적용상품지정</td>
	<td bgcolor="#FFFFFF">
		상품코드: <input type=text name=targetitemlist value="<%= targetitemlist %>" size=14 maxlength=14 readonly style='background-color:#E6E6E6;'>
		<input type="button" onClick="jsSearchItemID(frm, this.form.name,'targetitemlist')" value="검색" class='button'>
		(위탁 상품만 할인됨, 위탁 상품의 옵션 전부에 적용됨)
	</td>
</tr>
<tr height=30 id="trtargetbrandlist" style="display:<% if IsTargetBrandCoupon then %>block<% else %>none<% end if %>">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">적용브랜드지정</td>
	<td bgcolor="#FFFFFF">
		브 랜 드: <input type=text name=targetbrandlist value="<%= targetbrandlist %>" size=32 maxlength=32 readonly style='background-color:#E6E6E6;'>
		<input type="button" class="button" value="브랜드검색" onclick="jsSearchBrandID(this.form.name,'targetbrandlist');" >
		(위탁 브랜드만 할인됨)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">위탁상품 <font color=red>제외</font></td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="divexitemidlist" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				상품코드: <input type=text name=exitemidchoice value="" size=14 maxlength=14 readonly style='background-color:#E6E6E6;'>
				<input type="button" onClick="jsSearchItemID(frm, this.form.name,'exitemidchoice')" value="검색" class='button'>
				<input type="button" onClick="InsertExItemID()" value="추가" class='button'>
				(상품의 옵션 전부에 적용됨)
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
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">위탁브랜드 <font color=red>제외</font></td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="divexbrandidlist" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				브랜드: <input type=text name=exbrandidchoice value="" size=20 maxlength=32 readonly style='background-color:#E6E6E6;'>
				<input type="button" onClick="jsSearchBrandID(this.form.name,'exbrandidchoice')" value="검색" class='button'>
				<input type="button" onClick="InsertExBrandID()" value="추가" class='button'>
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
	<td bgcolor="<%= adminColor("tabletop") %>">대상고객</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isopenlistcoupon" value="N" <% if isopenlistcoupon="N" or isopenlistcoupon="" then Response.Write "checked" %>>전체고객
		<input type="radio" name="isopenlistcoupon" value="Y" <% if isopenlistcoupon="Y" then Response.Write "checked" %>>선택고객(지정고객)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">적용매장지정</td>
	<td bgcolor="#FFFFFF">
		<table border="0" id="div1" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td>
				<!-- 직영,가맹,해외점 = 1,3,7 -->
				<% drawSelectBoxOffShopdiv_off "shopidchoice",shopidchoice , "1" ,"","" %>
				<input type="button" onClick="AutoInsert()" value="추가" class='button'>
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
	<td bgcolor="<%= adminColor("tabletop") %>">사용혜택</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="couponvalue" value="<%= couponvalue %>" maxlength=7 size=10>
    	<% if coupontype="1" then %>
    		<input type="radio" name="coupontype" value="1" checked >%할인
    		<input type="radio" name="coupontype" value="2" >원할인
    	<% elseif coupontype="2" then %>
    		<input type="radio" name="coupontype" value="1" >%할인
    		<input type="radio" name="coupontype" value="2" checked >원할인
    	<% else %>
    		<input type="radio" name="coupontype" value="1" >%할인
    		<input type="radio" name="coupontype" value="2" checked >원할인
    	<% end if %>
		(금액 또는 % 할인)
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">유효기간</td>
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
	<td bgcolor="<%= adminColor("tabletop") %>">쿠폰발급마감일</td>
	<td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="openfinishdate" value="<%=left(openfinishdate,10)%>" size=10 readonly ><a href="javascript:calendarOpen(frm.openfinishdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    	<input type="text" name="openfinishdateTime" size="8" maxlength="8" class="text" value="<%=openfinishdateTime%>">
		(<%= Left(now(),10) %> 23:59:59)
	</td>
</tr>
<tr height=3 bgcolor="#FFFFFF"><td colspan=5></td></tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">기타코멘트</td>
	<td bgcolor="#FFFFFF"><textarea name="etcstr" cols=80 rows=8><%= etcstr %></textarea></td>
</tr>

<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" maxlength=7 size=10 <% if IsUsing="Y" or IsUsing="" then response.write " checked" %>>Y
		<input type="radio" name="isusing" value="N" maxlength=7 size=10>N
	</td>
</tr>
<tr height=30>
	<td colspan="2" align=center bgcolor="#FFFFFF">
		<input type=button value="저장" onClick="submitForm(frm);" class="button">
		<input type=button value="목록으로" onClick="location.href='/admin/offshop/bonuscoupon/couponlist.asp?menupos=<%=menupos%>';" class="button">
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