<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정기구독 대상자 발송
' History : 2016.06.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim itemid, itemoption, i, menupos, page, orderserial, userid, sendstatus
dim reserveDlvDate, reserveidx, reserveItemID, reserveItemOption, reserveItemName, regadminid, regdate
dim lastadminid, lastupdate, username, isusing, reloading, jukyogubun
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	reserveitemid = getNumeric(requestcheckvar(request("reserveitemid"),10))
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	itemoption = requestcheckvar(request("itemoption"),4)
	page = getNumeric(requestcheckvar(request("page"),10))
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	orderserial = requestcheckvar(request("orderserial"),11)
	username = requestcheckvar(request("username"),32)
	userid = requestcheckvar(request("userid"),32)
	isusing = requestcheckvar(request("isusing"),1)
	reloading = requestcheckvar(request("reloading"),2)
	sendstatus = requestcheckvar(request("sendstatus"),10)
	jukyogubun = requestcheckvar(request("jukyogubun"),16)

if reloading="" and isusing="" then isusing="Y"
if page="" then page=1

dim ouser
set ouser = new Citemstanding
	ouser.FPageSize = 300
	ouser.FCurrPage = page
	ouser.FRectItemID = itemid
	ouser.FRectreserveitemid = reserveitemid
	ouser.FRectitemoption = itemoption
	ouser.FRectreserveidx = reserveidx
	ouser.FRectorderserial = orderserial
	ouser.FRectusername = username
	ouser.FRectuserid = userid
	ouser.FRectisusing = isusing
	ouser.FRectsendstatus = sendstatus
	ouser.FRectjukyogubun = jukyogubun
	ouser.fitemstanding_user
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function chkAllchartItem() {
	if($("input[name='uidx']:first").attr("checked")=="checked") {
		$("input[name='uidx']").attr("checked",false);
	} else {
		$("input[name='uidx']").attr("checked","checked");
	}
}

function frmsubmit(page){
	if (frmstanding.itemid.value!=""){
		if (!IsDouble(frmstanding.itemid.value)){
			alert('판매용 상품코드는 숫자만 입력 가능합니다.');
			frmstanding.itemid.focus();
			return;
		}
	}
	if (frmstanding.reserveitemid.value!=""){
		if (!IsDouble(frmstanding.reserveitemid.value)){
			alert('배송 상품코드는 숫자만 입력 가능합니다.');
			frmstanding.reserveitemid.focus();
			return;
		}
	}
	frmstanding.itemoption.value=frmstanding.item_option_.value;

	frmstanding.page.value=page;
	frmstanding.submit();
}

// 기타출고 등록
function editstandinguser(uidx, editmode, reserveidx, itemid, itemoption){
	if (editmode=='RE' || editmode=='EDIT'){
		if (uidx==''){
			alert('일렬번호가 없습니다.');
			return false;
		}
	}else{
		if (itemid=='' || itemoption==''){
			alert('판매용상품코드와 판매용옵션코드를 우선 검색 하셔야 기타출고를 등록 하실수 있습니다.');
			return false;
		}
	}

	var editstandinguser = window.open('<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standinguser_edit.asp?uidx='+ uidx +'&editmode='+ editmode + '&reserveidx='+ reserveidx + '&itemid='+ itemid + '&itemoption='+ itemoption +'&menupos=<%= menupos %>','editstandinguser','width=800,height=600,scrollbars=yes,resizable=yes');
	editstandinguser.focus();
}

function savestandingsend() {
	var smsyn='';
	smsyn = 'N';
	if (frmstanding.smsyn.value=='Y'){
		if(confirm("고객님께 문자발송을 선택 하셨습니다. 문자를 발송 하시겠습니까?") == true) {
			smsyn = 'Y';
		}else{
			return false;
		}
	}

	var chk=0;
	$("form[name='frmstandinglist']").find("input[name='uidx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("발송처리 하실 항목을 선택해주세요.");
		return;
	}

	var uidx;
	for (i=0; i< frmstandinglist.uidx.length; i++){
		if (frmstandinglist.uidx[i].checked == true){
			uidx = frmstandinglist.uidx[i].value;

		    if ( !(eval("frmstandinglist.sendstatus_" + uidx).value=='0' || eval("frmstandinglist.sendstatus_" + uidx).value=='5') ){
		    	alert('선택하신 항목중에 발송대기나 재발송대기 상태가 아닌 항목이 있습니다.');
				eval("frmstandinglist.sendstatus_" + uidx).focus();
				return false;
		    }
	    }
	}

	if(confirm("선택하신 정기구독을 발송처리 하시겠습니까?")) {
		frmstandinglist.smsyn.value=smsyn;
		frmstandinglist.mode.value="savestandingsend";
		frmstandinglist.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standinguser_process.asp";
		frmstandinglist.target="";
		frmstandinglist.submit();
	}
}

// 엑셀 다운로드
function exceldownload(){
	frmstandinglist.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standinguser_excel.asp";
	frmstandinglist.target="view";
	frmstandinglist.submit();
}

//정기구독 발송 수기등록 가져오기
function regstandingusersudong(){
	var reserveidx='<%= reserveidx %>';

	if (frmstanding.item_option_.value==''){
		alert('옵션코드를 선택 하세요.');
		frmstanding.item_option_.focus();
		return false;
	}
	if (frmstanding.sendkey.value==''){
		alert('발송차수를 선택 하세요.');
		frmstanding.sendkey.focus();
		return false;
	}
	if (reserveidx==''){
		alert('발행회차 Vol.(번호)가 등록되지 않았습니다.');
		return false;
	}

	if(confirm("수기등록 차수1 대상자의 마지막(재발송포함) 배송지를 가져 옵니다.\n진행하시겠습니까?\n\n참고)신규추가,수정,재발송은 차수1 에 입력하세요.")) {
		frmstandinguserreg.itemid.value='<%= itemid %>';
		frmstandinguserreg.itemoption.value=frmstanding.item_option_.value;
		frmstandinguserreg.sendkey.value=frmstanding.sendkey.value;
		frmstandinguserreg.reserveidx.value=reserveidx;
		frmstandinguserreg.mode.value="standingusersudonginsert";
		frmstandinguserreg.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standinguser_process.asp";
		frmstandinguserreg.submit();
	}
}

</script>

<form name="frmstanding" method="get" action="" style="margin:0;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reloading" value="ON">

<Br>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 판매용상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=10 >
		<% if itemid<>"" then %>
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			&nbsp;* 판매용옵션코드 : <%= getOptionBoxHTML_FrontTypenew_optionisusingN_standingitem(itemid, itemoption, " onchange='frmsubmit("""");'") %>

			<% if itemoption<>"" then %>
				&nbsp;* 회차 : <% drawSelectBoxsendkey "reserveidx", reserveidx, itemid, itemoption, " onchange='frmsubmit("""");'" %>
			<% end if %>
		<% else %>
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			<input type="hidden" name="item_option_" value="<%= itemoption %>">
			<input type="hidden" name="reserveidx" value="<%= reserveidx %>">
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 배송상품코드 : <input type="text" name="reserveitemid" value="<%= reserveitemid %>" size=9 maxlength=10 >
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="14" maxlength="14">
		&nbsp;
		* 이름 : <input type="text" class="text" name="username" value="<%= username %>" size="7" maxlength="7">
		&nbsp;
		* 아이디 : <input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="20">
		&nbsp;
		* 사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""");'" %>
		&nbsp;
		* 상태 : <% drawSelectBoxsendstatus "sendstatus", sendstatus, " onchange='frmsubmit("""");'" %>
		&nbsp;
		* 적요 : <% drawSelectBoxjukyo "jukyogubun", jukyogubun, " onchange='frmsubmit("""");'" %>
	</td>
</tr>
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" onClick="exceldownload();" value="엑셀다운로드" class="button">
		<input type="button" onclick="editstandinguser('','SUDONG','<%= reserveidx %>','<%= itemid %>','<%= itemoption %>');" value="기타출고등록" class="button">
		<!--<input type="button" onclick="regstandingusersudong();" value="1차수기대상자가져오기" class="button">-->
	</td>
	<td align="right">
		문자발송 :
		<select name="smsyn">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
		<input type="button" onClick="savestandingsend();" value="선택발송처리" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
</form>

<form name="frmstandinglist" method="post" action="" style="margin:0;">
<input type="hidden" name="smsyn" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= itemid %>" >
<input type="hidden" name="reserveitemid" value="<%= reserveitemid %>" >
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="reserveidx" value="<%= reserveidx %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="username" value="<%= username %>">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="isusing" value="<%= isusing %>">
<input type="hidden" name="sendstatus" value="<%= sendstatus %>">
<input type="hidden" name="jukyogubun" value="<%= jukyogubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ouser.FtotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ouser.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=30><input type="button" value="전체" class="button" onClick="chkAllchartItem();"></td>
    <!--<td width=60>idx</td>-->
    <td width=60>발행회차<br>Vol.(번호)</td>
    <td width=60>배송<br>상품코드</td>
    <td width=50>배송<br>옵션코드</td>
    <td>배송상품명</td>
	<td width=70>적요</td>
    <td width=70>주문번호</td>
    <td width=40>수량</td>
    <td width=70>아이디</td>
    <td width=60>이름</td>
	<td width=60>상태</td>
	<td width=70>발송일</td>
	<td width=30>사용<br>여부</td>
    <td width=60>판매용<br>상품코드</td>
    <td width=50>판매용<br>옵션코드</td>
	<td width=60>비고</td>
</tr>

<% if ouser.FtotalCount>0 then %>
	<%
	for i=0 to ouser.FResultCount - 1
	%>
	<tr bgcolor="<%=chkIIF(ouser.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(ouser.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>'; align="center">
	    <td align="center"><input type="checkbox" name="uidx" value="<%= ouser.FItemList(i).fuidx %>"/></td>
	    <!--<td><%'= ouser.FItemList(i).fuidx %></td>-->
	    <td>
	    	<%= ouser.FItemList(i).freserveidx %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).freserveItemID %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).freserveItemOption %>
	    </td>
	    <td align="left">
	    	<%= ouser.FItemList(i).freserveItemname %>
	    </td>				
	    <td>
	    	<%= getjukyoname(ouser.FItemList(i).fjukyogubun) %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forderserial %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fitemno %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fuserid %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fusername %>
	    </td>
	    <td align="center">
	    	<input type="hidden" name="sendstatus_<%= ouser.FItemList(i).fuidx %>" class="text_ro" value="<%= ouser.FItemList(i).fsendstatus %>" />
	    	<font color="red"><%= getsendstatusname(ouser.FItemList(i).fsendstatus) %></font>
	    </td>
	    <td align="center">
	    	<%= left(ouser.FItemList(i).fsenddate,10) %>
	    	<Br><%= mid(ouser.FItemList(i).fsenddate,12,11) %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fisusing %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forgitemid %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forgitemoption %>
	    </td>
	    <td align="center">
	    	<% if ouser.FItemList(i).fsendstatus=0 or ouser.FItemList(i).fsendstatus=5 then %>
				<% if ouser.FItemList(i).fjukyogubun<>"ORDER" then %>
					<input type="button" onclick="editstandinguser('<%= ouser.FItemList(i).fuidx %>','EDIT','','','');" value="수정" class="button">
				<% end if %>
	    	<% end if %>

			<% if ouser.FItemList(i).fsendstatus=3 or ouser.FItemList(i).fsendstatus=7 then %>
	    		<input type="button" onclick="editstandinguser('<%= ouser.FItemList(i).fuidx %>','RE','','','');" value="재발송" class="button">
	    	<% end if %>
	    </td>
	</tr>
	<%
	Next
	%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if ouser.HasPreScroll then %>
			<a href="javascript:frmsubmit('<%= ouser.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ouser.StartScrollPage to ouser.FScrollCount + ouser.StartScrollPage - 1 %>
				<% if i>ouser.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ouser.HasNextScroll then %>
				<a href="javascript:frmsubmit('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>

</table>
</form>
<form name="frmstandinguserreg" method="POST" action="" style="margin:0;">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="reserveidx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="400" allowtransparency="true" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" allowtransparency="true" frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set ouser=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
