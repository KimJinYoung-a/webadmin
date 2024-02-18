<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim suplyer
suplyer = request("suplyer")

dim reguser, divcode,baljuname,regname,comment
dim shopid
reguser = session("ssBctid")
regname = session("ssBctCname")
divcode = request("divcode")
comment = html2db(request("comment"))


shopid = "10x10"
baljuname = "텐바이텐"


dim itemgubunarr, itemidadd, itemoptionarr
dim itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr

dim itemgubunarr2, itemidadd2, itemoptionarr2
dim itemnamearr2, itemoptionnamearr2
dim sellcasharr2, suplycasharr2, buycasharr2, itemnoarr2, designerarr2, mwdivarr2

dim itemgubunarr3, itemidadd3, itemoptionarr3
dim itemnamearr3, itemoptionnamearr3
dim sellcasharr3, suplycasharr3, buycasharr3, itemnoarr3, designerarr3, mwdivarr3

dim i,j,cnt,cnt2
dim scheduledt,executedt

itemgubunarr = request("itemgubunarr")
itemidadd	= request("itemidadd")
itemoptionarr = request("itemoptionarr")
itemnamearr		= request("itemnamearr")
itemoptionnamearr = request("itemoptionnamearr")
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
mwdivarr = request("mwdivarr")

itemgubunarr2 = request("itemgubunarr2")
itemidadd2	= request("itemidadd2")
itemoptionarr2 = request("itemoptionarr2")
itemnamearr2	= request("itemnamearr2")
itemoptionnamearr2 = request("itemoptionnamearr2")
sellcasharr2 = request("sellcasharr2")
suplycasharr2 = request("suplycasharr2")
buycasharr2 = request("buycasharr2")
itemnoarr2 = request("itemnoarr2")
designerarr2 = request("designerarr2")
mwdivarr2 = request("mwdivarr2")

itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
mwdivarr = split(mwdivarr,"|")

itemgubunarr2 = split(itemgubunarr2,"|")
itemidadd2	= split(itemidadd2,"|")
itemoptionarr2 = split(itemoptionarr2,"|")
itemnamearr2		= split(itemnamearr2,"|")
itemoptionnamearr2 = split(itemoptionnamearr2,"|")
sellcasharr2 = split(sellcasharr2,"|")
suplycasharr2 = split(suplycasharr2,"|")
buycasharr2 = split(buycasharr2,"|")
itemnoarr2 = split(itemnoarr2,"|")
designerarr2 = split(designerarr2,"|")
mwdivarr2 = split(mwdivarr2,"|")

scheduledt = request("scheduledt")
executedt = request("executedt")

cnt = uBound(itemidadd)
cnt2 = uBound(itemidadd2)


dim isPreExists

for j=0 to cnt2-1
	isPreExists = false
	for i=0 to cnt-1
		if (itemgubunarr(i)=itemgubunarr2(j)) and (itemidadd(i)=itemidadd2(j)) and (itemoptionarr(i)=itemoptionarr2(j)) then
			itemnoarr(i) = CStr(CLng(itemnoarr(i)) + CLng(itemnoarr2(j)))
			isPreExists = true
			exit for
		end if
	next

	if Not isPreExists then
		itemgubunarr3 = itemgubunarr3 + itemgubunarr2(j) + "|"
		itemidadd3	= itemidadd3 + itemidadd2(j) + "|"
		itemoptionarr3 = itemoptionarr3 + itemoptionarr2(j) + "|"
		itemnamearr3		= itemnamearr3 + itemnamearr2(j) + "|"
		itemoptionnamearr3  = itemoptionnamearr3 + itemoptionnamearr2(j) + "|"
		sellcasharr3 = sellcasharr3 + sellcasharr2(j) + "|"
		suplycasharr3 = suplycasharr3 + suplycasharr2(j) + "|"
		buycasharr3 = buycasharr3 + buycasharr2(j) + "|"
		itemnoarr3 = itemnoarr3 + itemnoarr2(j) + "|"
		designerarr3 = designerarr3 + designerarr2(j) + "|"
		mwdivarr3 = mwdivarr3 + mwdivarr2(j) + "|"
	end if
next

itemgubunarr2 = ""
itemidadd2	= ""
itemoptionarr2 = ""
itemnamearr2	= ""
itemoptionnamearr2 = ""
sellcasharr2 = ""
suplycasharr2 = ""
buycasharr2 = ""
itemnoarr2 = ""
designerarr2 = ""
mwdivarr2 = ""

for i=0 to cnt-1
	itemgubunarr2 = itemgubunarr2 + itemgubunarr(i) + "|"
	itemidadd2	= itemidadd2 + itemidadd(i) + "|"
	itemoptionarr2 = itemoptionarr2 + itemoptionarr(i) + "|"
	itemnamearr2	= itemnamearr2 + itemnamearr(i) + "|"
	itemoptionnamearr2 = itemoptionnamearr2 + itemoptionnamearr(i) + "|"
	sellcasharr2 = sellcasharr2 + sellcasharr(i) + "|"
	suplycasharr2 = suplycasharr2 + suplycasharr(i) + "|"
	buycasharr2 = buycasharr2 + buycasharr(i) + "|"
	itemnoarr2 = itemnoarr2 + itemnoarr(i) + "|"
	designerarr2 = designerarr2 + designerarr(i) + "|"
	mwdivarr2 = mwdivarr2 + mwdivarr(i) + "|"

next

itemgubunarr = itemgubunarr2 + itemgubunarr3
itemidadd	= itemidadd2 + itemidadd3
itemoptionarr = itemoptionarr2 + itemoptionarr3
itemnamearr	= itemnamearr2 + itemnamearr3
itemoptionnamearr = itemoptionnamearr2 + itemoptionnamearr3
sellcasharr = sellcasharr2 + sellcasharr3
suplycasharr = suplycasharr2 + suplycasharr3
buycasharr = buycasharr2 + buycasharr3
itemnoarr = itemnoarr2 + itemnoarr3
designerarr = designerarr2 + designerarr3
mwdivarr = mwdivarr2 + mwdivarr3

''디폴트 매입구분
dim sqlstr, maeipgubun
if (divcode="") and (suplyer<>"") then
	sqlstr = "select top 1 maeipdiv from [db_user].[dbo].tbl_user_c"
	sqlstr = sqlstr + " where userid='" + suplyer + "'"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		maeipgubun = rsget("maeipdiv")
	end if
	rsget.close

	if maeipgubun="M" then
		divcode = "001"
	elseif maeipgubun="W" then
		divcode = "002"
	end if
end if


dim BasicMonth, IsExpireEdit
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

IsExpireEdit = false

%>
<script language='javascript'>
function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,imwdiv){
	if (iidx!='0'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
		return;
	}

	frmMaster.itemgubunarr2.value = igubun;
	frmMaster.itemidadd2.value = iitemid;
	frmMaster.itemoptionarr2.value = iitemoption;
	frmMaster.sellcasharr2.value = isellcash;
	frmMaster.suplycasharr2.value = isuplycash;
	frmMaster.buycasharr2.value = ibuycash;
	frmMaster.itemnoarr2.value = iitemno;
	frmMaster.itemnamearr2.value = iitemname;
	frmMaster.itemoptionnamearr2.value = iitemoptionname;
	frmMaster.designerarr2.value = iitemdesigner;
	frmMaster.mwdivarr2.value = imwdiv;

	frmMaster.submit();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	popwin = window.open('popjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0' ,'ipgoinputadd','width=840,height=600,scrollbars=yes,resizable=no');
	popwin.focus();
}


function AddItemsNew(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	popwin = window.open('popjumunitemNew.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0' ,'ipgoinputaddnew','width=940,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItemsNewByBarCode(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	popwin = window.open('popjumunitemBybarcode.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0' ,'ipgoinputaddnewbarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ConFirmIpChulList(){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var pmwdiv = "";
	var frm;

	if (msfrm.scheduledt.value.length<1){
		alert('명세서일자를 입력해 주세요.');
		calendarOpen(msfrm.scheduledt);
		return;
	}

	if (msfrm.executedt.value.length<1){
		alert('실 입고일자를 입력해 주세요.');
		calendarOpen(msfrm.executedt);
		return;
	}

<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (msfrm.executedt.value>'<%= date() %>'){
		alert('입고일은 오늘날짜 보다 클수 없습니다.');
		return;
	}
<% end if %>
	if (msfrm.divcode[0].checked){
		upfrm.divcode.value = msfrm.divcode[0].value;
	}else if (msfrm.divcode[1].checked){
		upfrm.divcode.value = msfrm.divcode[1].value;
	}else if (msfrm.divcode[2].checked){
		upfrm.divcode.value = msfrm.divcode[2].value;
	}else if (msfrm.divcode[3].checked){
		upfrm.divcode.value = msfrm.divcode[3].value;
	}else{
		alert('매입구분을 선택해 주세요.');
		msfrm.divcode[0].focus();
		return;
	}

	if (upfrm.divcode.value=="001"){
		pmwdiv = "M";
	}else{
		pmwdiv = "W";
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";
	upfrm.mwdivarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				frm.itemno.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";

			//매입시 매입가,공급가 동일
			// upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.buycasharr.value = upfrm.buycasharr.value + frm.suplycash.value + "|";
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
			upfrm.mwdivarr.value = upfrm.mwdivarr.value + frm.mwdiv.value + "|";

			if ((frm.itemno.value!="0") && (pmwdiv!=frm.mwdiv.value)){
				if (!confirm(frm.itemid.value + '-매입 속성이 일치하지 않습니다.\r\n 계속 하시겠습니까?')){
					return;
				}
			}
		}
	}

	if (checkAvail3(msfrm.executedt.value) != true) {
		return;
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.scheduledt.value = msfrm.scheduledt.value;
		upfrm.executedt.value = msfrm.executedt.value;
		upfrm.comment.value = msfrm.comment.value;

		upfrm.submit();
	}
}

// 매월 5일까지 전월내역 수정가능
function checkAvail3(modiexecutedt) {
	var thisDate = "<%= Left(Now, 7) %>-01";
	var availDate = "<%= Left(Now, 7) %>-05";
	var nowdate = "<%= Left(now(),10) %>";
	var BasicMonth = "<%= BasicMonth %>";

	if (modiexecutedt < BasicMonth){
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert('변경불가\n\n입고일이 전월 이전날짜입니다.');
		return false;
		<% else %>
		alert('관리자권한\n\n입고일이 전월 이전날짜입니다.');
		<% end if %>
	}

	if (modiexecutedt < thisDate) {
		if (nowdate > availDate) {
			<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
			alert("변경불가\n\n매월 5일까지만 전월내역 변경가능합니다.");
			return false;
			<% else %>
			alert('관리자권한\n\n매월 5일까지만 전월내역 변경가능합니다.');
			<% end if %>
		}
	}

	return true;
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        	<font color="red"><strong>입고입력</strong></font>
	        </td>
	        <td align="right">

	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMaster" method="post" action="">
	<input type="hidden" name="mode" value="addipgo">
	<input type="hidden" name="chargeid" value="<%= session("ssBctId") %>">
	<input type="hidden" name="chargename" value="<%= session("ssBctCname") %>">
	<input type="hidden" name="vatcode" value="008">
	<input type="hidden" name="shopid" value="10x10">

	<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
	<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
	<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
	<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
	<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
	<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
	<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
	<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
	<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
	<input type="hidden" name="designerarr" value="<%= designerarr %>">
	<input type="hidden" name="mwdivarr" value="<%= mwdivarr %>">

	<input type="hidden" name="itemgubunarr2" value="">
	<input type="hidden" name="itemidadd2" value="">
	<input type="hidden" name="itemoptionarr2" value="">
	<input type="hidden" name="itemnamearr2" value="">
	<input type="hidden" name="itemoptionnamearr2" value="">
	<input type="hidden" name="sellcasharr2" value="">
	<input type="hidden" name="suplycasharr2" value="">
	<input type="hidden" name="buycasharr2" value="">
	<input type="hidden" name="itemnoarr2" value="">
	<input type="hidden" name="designerarr2" value="">
	<input type="hidden" name="mwdivarr2" value="">

	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">입출고코드</td>
		<td></td>
		<td width=100 bgcolor="<%= adminColor("tabletop") %>" >공급처</td>
		<% if suplyer<>"" then %>
		<input type=hidden name="suplyer" value="<%= suplyer %>">
		<td bgcolor="#FFFFFF"><%= suplyer %></td>
		<% else %>
		<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "suplyer", suplyer %></td>
		<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >명세서일자</td>
		<td><input type="text" name="scheduledt" value="<%= scheduledt %>" size="10" maxlength=10 readonly><a href="javascript:calendarOpen(frmMaster.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a></td>
		<td bgcolor="<%= adminColor("tabletop") %>" >입고일자</td>
		<td><input type="text" name="executedt" value="<%= executedt %>" size="10" maxlength=10 readonly><a href="javascript:calendarOpen(frmMaster.executedt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (재고와 관련 있음)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">매입구분</td>
		<td>
			<input type="radio" name="divcode" value="001" <% if divcode="001" then response.write "checked" %> >매입
			<input type="radio" name="divcode" value="002" <% if divcode="002" then response.write "checked" %> >위탁
			<input type="radio" name="divcode" value="801" <% if divcode="801" then response.write "checked" else response.write "disabled" end if %> >OFF용매입
			<input type="radio" name="divcode" value="802" <% if divcode="802" then response.write "checked" else response.write "disabled" end if %> >OFF용위탁
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">입고자</td>
		<td><%= session("ssBctCname") %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >기타사항</td>
		<td colspan="3"><textarea name="comment" cols=80 rows=6></textarea></td>
	</tr>
	</form>
</table>




<%
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
mwdivarr = split(mwdivarr,"|")

cnt = ubound(itemidadd)

dim selltotal, suplytotal
selltotal =0
suplytotal =0
%>
<br>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        	<font color="red"><strong>상세내역</strong></font>
	        	&nbsp;&nbsp;
	        	<font color="<%= mwdivColor("M") %>">매입</font>&nbsp;
	        	<font color="<%= mwdivColor("W") %>">위탁</font>&nbsp;
	        	<font color="<%= mwdivColor("U") %>">업체배송</font>
	        </td>
	        <td align="right">
	        	총건수:  <% if (cnt = -1) then %>0<% else %><%= cnt %><% end if %>
	        	&nbsp;&nbsp;
				<!--<input type=button value=" 상품추가(old)" onClick="AddItems(frmMaster)"> -->
				<input type=button value=" 상품추가(바코드) " onClick="AddItemsNewByBarCode(frmMaster)">
				<input type=button value=" 상품추가" onClick="AddItemsNew(frmMaster)">
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90">바코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="80">판매가</td>
		<td width="80">매입가</td>
		<td width="60">마진</td>
		<td width="60">수량</td>
	</tr>
	<% for i=0 to cnt-1 %>
	<%
	selltotal  = selltotal + sellcasharr(i) * itemnoarr(i)
	suplytotal = suplytotal + suplycasharr(i) * itemnoarr(i)
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="">
	<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
	<input type="hidden" name="itemid" value="<%= itemidadd(i) %>">
	<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
	<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
	<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
	<input type="hidden" name="buycash" value="<%= buycasharr(i) %>">
	<input type="hidden" name="mwdiv" value="<%= mwdivarr(i) %>">

	<tr bgcolor="#FFFFFF">
		<td>
			<font color="<%= mwdivColor(mwdivarr(i)) %>"><%= itemgubunarr(i) %>-<%= CHKIIF(itemidadd(i)>=1000000,format00(8,itemidadd(i)),format00(6,itemidadd(i))) %>-<%= itemoptionarr(i) %></font>
		</td>
		<td ><%= itemnamearr(i) %></td>
		<td ><%= itemoptionnamearr(i) %></td>
		<td align=right><%= FormatNumber(sellcasharr(i),0) %></td>
		<td align=right>
		<input type="text" name="suplycash" value="<%= suplycasharr(i) %>" size=7 maxlength=7>
		</td>
		<td align=center>
		<% if sellcasharr(i)<>0 then %>
			<%= 100-CLng(suplycasharr(i)/sellcasharr(i)*100*100)/100 %>%
		<% end if %>
		</td>
		<td ><input type="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
	</tr>
	</form>
	<% next %>

	<% if (cnt>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan="2" align="center">
		<td align=right><%= formatNumber(selltotal,0) %></td>
		<td align=right><%= formatNumber(suplytotal,0) %></td>
		<td></td>
		<td></td>
	</tr>
	<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if (cnt>0) then %>
        	<input type="button" value="내역확정" onclick="ConFirmIpChulList()">
        	<% else %>
        	&nbsp;
        	<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
'// 등록자 아이디 + 시간을 가지고 중복입력 체크
dim uniqregdate : uniqregdate = getDatabaseTime()
%>

<form name="frmArrupdate" method="post" action="ipchuledit_process.asp">
<input type="hidden" name="mode" value="addipgo">
<input type="hidden" name="scheduledt" value="">
<input type="hidden" name="executedt" value="">

<input type="hidden" name="baljuid" value="<%= shopid %>">
<input type="hidden" name="targetid" value="<%= suplyer %>">
<input type="hidden" name="reguser" value="<%= reguser %>">
<input type="hidden" name="uniqregdate" value="<%= uniqregdate %>">
<input type="hidden" name="divcode" value="">
<input type="hidden" name="vatinclude" value="Y">
<input type="hidden" name="comment" value="">
<input type="hidden" name="regname" value="<%= regname %>">
<input type="hidden" name="baljuname" value="<%= baljuname %>">

<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
</form>
<% if (request("idx") <> "") then %>
<script>
// window.open('poplimitcheckipgo.asp?idx=<%= request("idx") %>', '_blank', 'width=800,height=450,scrollbars=yes,resizable=yes');
window.open('poplimitcheckipgoNew.asp?idx=<%= request("idx") %>', '_blank', 'width=800,height=450,scrollbars=yes,resizable=yes');
location.replace('/admin/newstorage/ipgolist.asp?menupos=539');
</script>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
