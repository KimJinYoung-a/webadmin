<%@ language=vbscript %>
<% option explicit %>
<!--
###########################################

2007-12-07,김정인

설명:
	1. 현재상태(statecd)별 진행단계수정 ---  "주문확인"후 "주문접수" 단계로 수정불가
	2. "출고 완료"시 입력정확성 검사
###########################################

-->
<!-- #include virtual="/designer/incSessiondesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->

<%
dim idx, isfixed, opage, ourl
idx = request("idx")
opage = request("opage")
ourl = request("ourl")
if idx="" then idx=0

dim ojumunmaster, ojumundetail

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster
isfixed = ojumunmaster.FOneItem.IsFixed

set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.GetOrderSheetDetail

dim yyyymmdd
yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)
%>
<script language='javascript'>

function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popjumunsheet.asp?idx=' + v + '&itype=' + itype,'popjumunsheet','width=760,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

<% if ojumunmaster.FOneItem.FStatecd="0" then %>
var jumunwait = true;
<% else %>
var jumunwait = false;
<% end if %>

<% if ojumunmaster.FOneItem.FStatecd="1" then %>
var jumunconfirm = true;
<% else %>
var jumunconfirm = false;
<% end if %>

function AddItems(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid  = frm.shopid.value;
	popwin = window.open('popjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid,'shopjumunitem','width=800,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ModiThis(frm){
	if (jumunconfirm!=true){
		alert('주문확인 상태가 아니면 수정하실 수 없습니다.');
		return;
	}
	var compno = eval(frm.baljuitemno.value)>eval(frm.realitemno.value)?true:false


	if(compno){
		if(frm.dtstat.value=='ipt'){//직적입력
			if (frm.comment.value==''){
				alert('확정수량이 적습니다.\사유를 입력해주세요');
				frm.comment.focus();
				return false;
			}
		}else if(frm.dtstat.value=='sso'){//일시품절
			if(frm.comment.value==''){
				alert('확정수량이 적습니다.\재입고 예정일을 입력해주세요');
				frm.comment.focus();
				return false ;
			}
		}else if(frm.dtstat==''){

		}

	}else{
		//frm.comdiv.style.display='none';
		//frm.comdiv.value='';
	}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}



function ModiState(frm){

	if (frm.scheduleipgodate.value.length<1){
		alert('입고예정일을 입력하세요.');
		frm.scheduleipgodate.focus();
		return;
	}
	var statval;

	for(var i=0;i<frm.statecd.length;i++){
		if (frm.statecd[i].checked){
			statval= frm.statecd[i].value;
		}
	}

	if (statval==''){
		alert('오류');
		return;
	}
	if (statval==7){
		if(frm.songjangdiv.value==''){
			alert('택배사를 선택해 주세요');
			return;
		}
		if(fntrim(frm.songjangno.value)==''){
			alert('송장번호를 입력해주세요');
			return;
		}
		if(fntrim(frm.beasongdate.value)==''){
			alert('출고일을 입력해주세요');
			return;
		}
	}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modistate";
		frm.submit();
	}
}
// 주문확인
function ModiStateConfirm(frm){


	if (frm.scheduleipgodate.value.length<1){
		alert('입고예정일을 입력하세요.');
		frm.scheduleipgodate.focus();
		return;
	}

	var ret = confirm('주문을 확인하시겠습니까?');

	if (ret){
		frm.statecd.value=1;
		frm.mode.value="modistate";
		frm.submit();
	}
}



//공백제거
function fntrim(str){
	var index, len
	while(true){
		index = str.indexOf(" ")
		if (index == -1) break;
		len = str.length;
		str = str.substring(0, index) + str.substring((index+1),len);
	}
return str;
}

//확정수량&사유별 액션
function chkRealItemNo(tn){
	var frm = eval("frmBuyPrc_"+ tn);
	var t = frm.baljuitemno;
	var v= frm.realitemno;

	if (isNaN(v.value)||v.value.length<1){
		v.value=0;
	}else{
		v.value=parseInt(v.value);
	}

	var seldiv = eval("seldiv" + tn);

	if(parseInt(t.value) > v.value){
		if (frm.dtstat!=''){
			seldiv.innerHTML='<select name="dtstat" onchange="fnselcom(this.value,' + tn +');"><option value="ipt">직접입력</option><option value="so">단종</option><option value="sso">일시품절</option></select><br>';
			fnselcom('ipt',tn);
		}else{
			seldiv.innerHTML='';
			fnselcom('',tn);
		}
	}else{
		seldiv.innerHTML='<input type="text" name="comment" value=""  size="8" maxlength="10">';
		fnselcom('',tn);
	}

}
//사유별 표시
function fnselcom(val,tn){
	var comdiv = eval("comdiv" + tn);
	if(val=='ipt'){
		comdiv.innerHTML='<input type="text" name="comment" value=""  size="10" maxlength="10">';
	}else if(val=='sso'){
		comdiv.innerHTML ='<input type="text" name="comment" value="" size="10" readonly ><a href="javascript:calendarOpen(frmBuyPrc_'+tn+'.comment);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>';
	}else{
		comdiv.innerHTML ='';
	}
}
function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	frmadd.itemgubunarr.value = igubun;
	frmadd.itemarr.value = iitemid;
	frmadd.itemoptionarr.value = iitemoption;
	frmadd.sellcasharr.value = isellcash;
	frmadd.suplycasharr.value = isuplycash;
	frmadd.buycasharr.value = ibuycash;
	frmadd.itemnoarr.value = iitemno;

	frmadd.submit();
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
	        <font color="red"><strong>주문정보</strong></font>
	        &nbsp;
	        <img src="/images/icon_num01.gif" align="absbottom">
	        <font color="red">주문확인 선택후 저장하시면, 상품상세내역이 표시됩니다.</font>

        </td>
        <td align="right">

        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMaster" method="post" action="jumuninputedit_process.asp">
	<input type=hidden name="mode" value="">
	<input type=hidden name="opage" value="<%= opage %>">
	<input type=hidden name="ourl" value="<%= ourl %>">
	<input type=hidden name="masteridx" value="<%= idx %>">
	<input type=hidden name="shopid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>">
    <tr bgcolor="#FFFFFF">
    	<td width="110" align="center" bgcolor="<%= adminColor("tabletop") %>">발주(주문)자</td>
		<td width="360"><%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)</td>
		<td width="110" align="center" bgcolor="<%= adminColor("tabletop") %>">공급자</td>
		<td>
			<input type=hidden name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
			<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">주문일시</td>
		<td><%= ojumunmaster.FOneItem.Fregdate %></td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">입고요청일</td>
		<td><%= yyyymmdd %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">현재상태</td>
		<td>
			<% if ojumunmaster.FOneItem.FStatecd>7 then %>
			<%= ojumunmaster.FOneItem.GetStateName %>
			<% else %>
				<% if ojumunmaster.FOneItem.FStatecd < 1 then %>
					<%= ojumunmaster.FOneItem.GetStateName %>
					<input type="hidden" name="statecd" value="0" >
				<% else %>
					<input type=radio name="statecd" value="1" <% if ojumunmaster.FOneItem.FStatecd="1" then response.write "checked" %> >주문확인
				<input type=radio name="statecd" value="7" <% if ojumunmaster.FOneItem.FStatecd="7" then response.write "checked" %> >출고완료
				<% end if %>

			<% end if %>

		</td>

		<td align="center" bgcolor="<%= adminColor("tabletop") %>">입고예정일</td>
		<td>
			<input type=text name="scheduleipgodate" value="<%= ojumunmaster.FOneItem.Fscheduleipgodate %>" size=10 readonly >
			<a href="javascript:calendarOpen(frmMaster.scheduleipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		</td>
	</tr>
	<% if ojumunmaster.FOneItem.FStatecd>="1" then %>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">운송장입력</td>
		<td>
			택배사 : <% drawSelectBoxDeliverCompany "songjangdiv", trim(ojumunmaster.FOneItem.Fsongjangdiv) %>
			운송장번호 : <input type="text" name="songjangno" size=12 maxlength=16 value="<%= ojumunmaster.FOneItem.Fsongjangno %>" >
			<input type=hidden name="songjangname" value="">
		</td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">출고일</td>
		<td>
			<input type=text name="beasongdate" value="<%= ojumunmaster.FOneItem.Fbeasongdate %>" size=10 readonly >
			<a href="javascript:calendarOpen(frmMaster.beasongdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		</td>

	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">매입구분</td>
		<td colspan="3">
			<%= ojumunmaster.FOneItem.getdivcodename %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">소비자가합계(주문)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">공급가합계(주문)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunbuycash,0) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">소비자가합계(확정)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">공급가합계(확정)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %></b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">기타요청사항</td>
		<td colspan="3">
			<%= nl2br(ojumunmaster.FOneItem.FComment) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td colspan="3">
			<textarea name=replycomment cols=80 rows=6><%= ojumunmaster.FOneItem.FReplyComment %></textarea>
		</td>
	</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if ojumunmaster.FOneItem.FStatecd=0 Then %>
				<input type="button" value="주문확인" class="button" onclick="ModiStateConfirm(frmMaster)">
			<% elseif ojumunmaster.FOneItem.FStatecd=1 then %>
				<input type=button value="기본내용수정" onclick="ModiState(frmMaster)">
			<% elseif ojumunmaster.FOneItem.FStatecd>7 then %>
				&nbsp;
			<% else %>
				'1이상 7이하
			<% end if %>

        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<br>
<%

dim i,selltotal, suplytotal, buytotal
dim selltotalfix, suplytotalfix, buytotalfix
selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
%>


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
	        &nbsp;
	        <% if ojumunmaster.FOneItem.FStatecd>="1" then %>
	        <a href="javascript:PopIpgoSheet('<%= ojumunmaster.FOneItem.FIdx %>','2');"><img src="/images/icon_print_ipgosheet.gif" align="absbottom" border="0"></a>
	        &nbsp;
	        <img src="/images/icon_num02.gif" align="absbottom">
	        <font color="red">수량부족시 확정수량을 수정하시고, 거래명세서를 꼭 동봉하여 보내주세요.</font>
	        <% end if %>
        	<!--
			<input type=button value="상품추가" onclick="AddItems(frmMaster)">

			&nbsp;&nbsp;&nbsp;
			<input type=button value="전체저장" onclick="ModiArr(frmMaster)">
			-->
        </td>
        <td align="right">
        	<% if ojumunmaster.FOneItem.FStatecd>="1" then %>
        	총건수:  <%= ojumundetail.FResultCount %>
        	<% end if %>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90">바코드</td>
		<td width="50">이미지</td>
		<td width="80">브랜드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="70">판매가</td>
		<td width="70">매입가</td>
		<td width="50">주문수량</td>
		<td width="50">확정수량</td>
		<% if isfixed then %>
		<td width="80">비고</td>
		<% else %>
		<td width="80">비고</td>
		<td width="60">수정</td>
		<% end if %>
	</tr>
<% if ojumunmaster.FOneItem.FStatecd>="1" then %>
	<% for i=0 to ojumundetail.FResultCount-1 %>
	<%
	selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
	suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
	buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno

	selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
	suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
	buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="jumuninputedit_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx%>">
	<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
	<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
	<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
	<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
	<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).FSuplycash %>">
	<input type="hidden" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>">
	<input type="hidden" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ojumundetail.FItemList(i).FItemGubun %>-<%= format00(6,ojumundetail.FItemList(i).FItemID) %>-<%= ojumundetail.FItemList(i).Fitemoption %></td>
		<td><img src="<%= ojumundetail.FItemList(i).Fsmallimage %>" border="0"></td>
		<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
		<td><%= ojumundetail.FItemList(i).Fitemname %></td>
		<td><%= ojumundetail.FItemList(i).Fitemoptionname %></td>
		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fbuycash,0) %></td>
		<td><%= ojumundetail.FItemList(i).Fbaljuitemno %></td>
		<td><input type="text" name="realitemno" value="<%= ojumundetail.FItemList(i).Frealitemno %>"  size="4" maxlength="4" onKeyup="chkRealItemNo(<%= i %>);"></td>
		<% if isfixed then %>
		<td><%= ojumundetail.FItemList(i).Fcomment %></td>
		<% else %>
		<td>
			<span id="seldiv<%=i%>">
				<% if ojumundetail.FItemList(i).FDetail_status<>"" then %>
					<select name="dtstat" onchange="fnselcom(this.value,<%= i %>);">
						<option value="so" <% if ojumundetail.FItemList(i).FDetail_status="단종" then response.write "selected" %>>단종</option>
						<option value="ipt" <% if ojumundetail.FItemList(i).FDetail_status="직접입력" then response.write "selected" %>>직접입력</option>
						<option value="sso" <% if ojumundetail.FItemList(i).FDetail_status="일시품절" then response.write "selected" %>>일시품절</option>
					</select><br>
				<% end if %>
			</span>
			<span id="comdiv<%=i%>">
				<% if (ojumundetail.FItemList(i).FDetail_status="직접입력") then %>
					<input type="text" name="comment" value="<%= ojumundetail.FItemList(i).Fdetail_description %>"  size="10" maxlength="10">
				<% Elseif (ojumundetail.FItemList(i).FDetail_status="일시품절") then %>
					<input type="text" name="comment" value="<%= ojumundetail.FItemList(i).Fdetail_description %>" size="10" readonly ><a href="javascript:calendarOpen(frmBuyPrc_<%=i%>.comment);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
				<% end if %></span>
		</td>
		<td><input type="image" src="/images/icon_modify.gif" width="45" height="20" border="0" onclick="ModiThis(frmBuyPrc_<%= i %>) ;return false;"></td>
		<% end if %>
	</tr>
	</form>
	<% next %>

<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="11"><font color="red">주문확인 선택후 저장하시면, 상품상세내역이 표시됩니다.</font></td>
	</tr>
<% end if %>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left">
        	주문 소비자가계 : <b><%= formatNumber(selltotal,0) %></b>
        	&nbsp;&nbsp;
        	주문 공급가계 : <b><%= formatNumber(buytotal,0) %></b>
        </td>
        <td valign="bottom" align="right">
        	확정 소비자가계 : <b><%= formatNumber(selltotalfix,0) %></b>
        	&nbsp;&nbsp;
        	확정 공급가계 : <b><%= formatNumber(buytotalfix,0) %></b>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->










<% if ojumunmaster.FOneItem.FStatecd="0" then %>
<script language='javascript'>
alert('현재상태를 주문확인으로 전환하시면 상품상세내역이 표시됩니다.');
</script>
<% end if %>

<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<form name="frmadd" method=post action="jumuninputedit_process.asp">
<input type=hidden name="mode" value="shopjumunitemaddarr">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="itemgubunarr" value="">
<input type=hidden name="itemarr" value="">
<input type=hidden name="itemoptionarr" value="">
<input type=hidden name="sellcasharr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="itemnoarr" value="">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->