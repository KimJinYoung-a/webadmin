<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품수정
' Hieditor : 서동석 생성
'			 2021.03.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, menupos, i, defaultcheck
	itemid = getNumeric(requestCheckVar(request("itemid"),10))
	menupos = requestCheckVar(request("menupos"),10)

if itemid = "" then
	response.write "<script>"
	response.write "	alert('상품코드가 없습니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if

'####### 상품고시법에 의한 빈값유무체크
If IsNumeric(itemid) = false Then
	response.write "<script>"
	response.write "	alert('잘못된 상품코드입니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF
Dim vQuery, vIsOK
''vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
''rsget.open vQuery,dbget,1
''2015/06/18
vQuery = "[db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check]('" & itemid & "')"
rsget.CursorLocation = adUseClient
rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
If Not rsget.Eof Then
	vIsOK = rsget(0)
Else
	vIsOK = "x"
End IF
rsget.close()
'rw vIsOK
'####### 상품고시법에 의한 빈값유무체크

dim oitem
set oitem = new CItem
	oitem.FRectItemID = itemid
	oitem.FRectSellReserve ="Y"
	'if itemid<>"" then //상품번호 빈값 체크 제대로 되는지 확인위해 주석처리 2014.03.11 정윤정
		oitem.GetOneItem
	'end if

If oitem.FTotalCount < 1 Then
	response.write "<script type='text/javascript'>"
	response.write "	alert('존재하지 않는 상품코드입니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

dim oitemoption
set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

dim oitemrackoption
set oitemrackoption = new CItemOption
	oitemrackoption.FPageSize = 50
	oitemrackoption.FCurrPage = 1
	oitemrackoption.frectitemgubun = "10"
	oitemrackoption.FRectItemID = itemid

	if itemid<>"" then
		oitemrackoption.GetItemrackcodeInfo
	end if

''브랜드 랙코드
dim sqlStr, prtidx
dim objCmd, returnValue
if (itemid<>"") and (oitem.FResultCount>0) then
    sqlStr = "select prtidx from [db_user].[dbo].tbl_user_c "
    sqlStr = sqlStr & " where userid='" & oitem.FOneItem.FMakerid & "'"
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        prtidx = rsget("prtidx")

        prtidx = format00(4,prtidx)
    end if
    rsget.close

'### 오픈예약 조건 체크
if oitem.FOneItem.Fdeliverytype = "1" or  oitem.FOneItem.Fdeliverytype = "4" then '텐바이텐 배송일때 재고여부 확인
set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_chkStock]("&itemid&")}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
Set objCmd = nothing
else
	returnValue = 1
end if
end if
'response.write "type="&oitem.FOneItem.Fdeliverytype&"rV="&returnValue
if oitem.FOneItem.Fsellreservedate ="" or isnull(oitem.FOneItem.Fsellreservedate) then
	oitem.FOneItem.Fsellreservedate=now()
	defaultcheck=false
else
	defaultcheck=true
end if

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/datetime.js?v=1.0"></script>
<script type="text/javascript">

//한정 or 비한정 Radio버튼 클릭시
function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && (e.name.substring(0,"optremainno".length) == "optremainno")) {
				e.disabled = (comp.value=="N");
		  }
  	}

    frm.recalcuLimit.disabled = (comp.value=="N");

    if (comp.value=="N"){ //비한정
        resetLimit2Zero();
        document.all.dvDisp.style.display = "none";
        frm.limitdispyn[0].checked = false;
        frm.limitdispyn[1].checked = true;
    }else{ //한정
        resetLimit();
        document.all.dvDisp.style.display = "";
    }
}

//한정수량 재설정
function resetLimit(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    //Enable 인 경우만
		  	    if (!e.disabled){
		  	        //현재 재고의 98%로 설정 (재고가 10개 이상인 경우만) 내림(97% -> 98%, 2014-07-25)
		  	        //if (e.getAttribute("dumistock")>=10){
		  	          //  e.value = parseInt(e.getAttribute("dumistock")*0.98);
		  	       // }

					//2016.04.08 현재 재고 수량과 동일하게 재 설정		'/2016.04.08 정윤정 추가
					//한정비교재고가 0보다 작을경우 0을 셋팅	'/2016.04.20 한용민 추가
					if ( parseInt(e.getAttribute("dumistock"))<0 ){
						e.value = 0;
					}else{
						e.value = parseInt(e.getAttribute("dumistock"));
					}
		  	    }
		  	}
		}
  	}
}

//한정수량 0으로 Setting
function resetLimit2Zero(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    e.value = 0;
		  	}
		}
  	}
}

function SaveItem(frm){
	var obj, subobj;
    var i, optdanjongyn, optisusing0, optisusing1;

	if ((frm.itemrackcode.value.length > 0) && (frm.itemrackcode.value.length != 4) && (frm.itemrackcode.value.length != 8)){
		alert('상품 랙코드는 4자리 또는 8자리로 고정되어있습니다.');
		frm.itemrackcode.focus();
		return;
	}

    if ((frm.subitemrackcode.value.length > 0) && (frm.subitemrackcode.value.length != 4) && (frm.subitemrackcode.value.length != 8)){
		alert('상품 보조랙코드는 4자리 또는 8자리로 고정되어있습니다.');
		frm.subitemrackcode.focus();
		return;
	}

	<% if oitem.FResultCount>0 then %>
	    <% if Not oitem.FOneItem.IsUpchebeasong then %>
	    // 주석처리 - 2014.04.01 정윤정
	    //판매 N 인경우 단종품절 또는 MD품절로 설정 해야함.
	   // if ((frm.sellyn[2].checked)&&!((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked))){
	     //   alert('판매 중지 상품인경우 재고부족,단종품절 또는 MD품절로 설정하셔야 합니다.');
	       // frm.danjongyn[2].focus();
	        //return;
	    //}

	    //재고부족,단종설정은 한정판매인경우만 가능함 (판매시만으로 변경)
		if ((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked)){
			if ((frm.sellyn[0].checked)&&(!frm.limityn[0].checked)){
				alert('판매중이고, 한정 판매인 경우만 재고부족,단종품절, MD품절로 설정 할 수 있습니다.');
				frm.limityn[0].focus();
				return;
			}
		}
    	<% if oitemoption.FResultCount > 0 then %>
        for (i = 0; ;i++) {
            optisusing0 = document.getElementById('optisusing0_' + i);
            optisusing1 = document.getElementById('optisusing1_' + i);
            optdanjongyn = document.getElementById('optdanjongyn_' + i);

            if (optdanjongyn == undefined) { break; }
            if (optisusing1.checked == true) { continue; }

            if ((optdanjongyn.value == 'S') || (optdanjongyn.value == 'Y') || (optdanjongyn.value == 'M')) {
	    		if ((frm.sellyn[0].checked)&&(!frm.limityn[0].checked)){
		    		alert('판매중이고, 한정 판매인 경우만 재고부족,단종품절, MD품절로 설정 할 수 있습니다.');
			    	frm.limityn[0].focus();
				    return;
    			}
            }
        }
    	<% end if %>
    	<% end if %>
	<% end if %>

	//사용안함이나 전시하는경우
	if ((frm.isusing[1].checked)&&(frm.sellyn[0].checked)){
        alert('사용 중지 상품은 판매로 설정 불가합니다.');
        frm.sellyn[2].focus();
        return;
    }

	//오픈예약
	if(typeof(frm.chkSR)=="object"){
		if(frm.chkSR.checked){
	    if(frm.dSR.value==""){
	    	alert("오픈예약이 설정되어있습니다. 날짜를 입력해주세요");
	    	frm.dSR.focus();
	    	return;
	    }

		if(toDate(frm.dSR.value+" "+frm.settime.value+":00:00") <= toDate("<%=date() &" "& Num2Str(hour(now()),2,"0","R") & ":00:00"%>")){
		 	alert("날짜/시간 선택이 잘못되었습니다.\n\n※ 예약 날짜와 시간을 현재 이후로 선택해주세요.");
			frm.dSR.focus();
			return;
		}

	    if(frm.sellyn[0].checked){
		 	 	if(confirm(frm.dSR.value+"로 오픈예약된 상품입니다. 판매중으로 상태 변경하시면, 상품오픈예약설정은 취소됩니다. 계속하시겠습니까? ")){
		 	 		frm.dSR.value = "";
		 	 		frm.chkSR.checked= false;
		 	 	}else{
		 	 		frm.sellyn[0].focus();
		 	 		return;
		 	 	}
	 		}

		 	if(frm.sellyn[1].checked){
			 	 	if(confirm(frm.dSR.value+"로 오픈예약된 상품입니다. 일시품절로 상태 변경하시면, 상품오픈예약설정은 취소됩니다. 계속하시겠습니까? ")){
			 	 		frm.dSR.value = "";
			 	 		frm.chkSR.checked= false;
			 	 	}else{
			 	 		frm.sellyn[1].focus();
			 	 		return;
			 	 	}
		 	}

	 		if(frm.chkSRC.value==0){
	 			alert("텐바이텐 배송일 경우, 입고 확인 후 오픈예약이 가능합니다.");
	 			frm.chkSR.focus();
	 			return;
	 		}
	   }
	}

	frm.itemoptionarr.value = "";
	//옵션 한정 남은 수량
	frm.optremainnoarr.value = "";
	frm.optrackcodearr.value = "";
    frm.suboptrackcodearr.value = "";
	//옵션 사용 여부
	frm.optisusingarr.value = "";
    frm.optdanjongynarr.value = '';

    var option_isusing_count = 0;
	var curritemoption;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
				curritemoption = e.id;
		  	    //숫자만 가능
		  	    if (!IsDigit(e.value)){
		  	        alert('한정 수량은 숫자만 가능합니다.');
		  	        e.select();
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + curritemoption + "," ;
				frm.optremainnoarr.value = frm.optremainnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                } else {
					obj = document.getElementById("optrackcode" + curritemoption);
                    subobj = document.getElementById("suboptrackcode" + curritemoption);
					if ((obj.value.length > 0) && (obj.value.length != 4) && (obj.value.length != 8)){
						alert('상품 옵션 랙코드는 4자리 또는 8자리로 고정되어있습니다.');
						obj.focus();
						return;
					}
                    if ((subobj.value.length > 0) && (subobj.value.length != 4) && (subobj.value.length != 8)){
						alert('상품 옵션 보조 랙코드는 4자리 또는 8자리로 고정되어있습니다.');
						subobj.focus();
						return;
					}
					frm.optrackcodearr.value = frm.optrackcodearr.value + obj.value + "," ;
                    frm.suboptrackcodearr.value = frm.suboptrackcodearr.value + subobj.value + "," ;
				}
		  	}

            //옵션 사용여부
			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		} else if (e.type == 'select-one') {
			if ((e.name.substring(0,"optdanjongyn".length)) == "optdanjongyn") {
				frm.optdanjongynarr.value = frm.optdanjongynarr.value + e.value + "," ;
			}
        }
  	}

    if (option_isusing_count < 1) {
        alert("모든 옵션을 사용안함으로 할수 없습니다. 상품정보를 사용안함으로 변경하거나, 전시안함 변경하세요.");
        return;
    }

	<%
	If vIsOK = "x" Then
		If oitem.FOneItem.FSellYn <> "Y" Then
	%>
			if(frm.sellyn[0].checked)
			{
				var ret = confirm('상품고시내용이 모두 입력되어 있지 않은 상태입니다.\n그래도 판매함으로 저장 하시겠습니까?');
			}
			else
			{
				var ret = confirm('저장 하시겠습니까?');
			}
	<%	Else %>
			var ret = confirm('저장 하시겠습니까?');
	<%	End If
	Else
	%>
		var ret = confirm('저장 하시겠습니까?');
	<% End If %>

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + iid,'popitemoptionedit','width=1200 height=800 scrollbars=yes resizable=yes');
	popwin.focus();
}

function jsPopItemHistory(itemid){
	var popwin = window.open('/common/pop_itemhistory.asp?itemid=' + itemid,'jsPopItemHistory','width=1400 height=800 scrollbars=yes resizable=yes');
	popwin.focus();
}

//달력
function jsPopCal(sName){
 if(!document.all.chkSR.checked){
 	 document.all.chkSR.checked= true;
 	}
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

//오픈예약
function jsChkSellReserve(){
	if(!document.all.chkSR.checked){
		document.all.dSR.value = "";
	}
}

function CloseWindow() {
    window.close();
}

function ReloadWindow() {
    document.location.reload();
}

window.resizeTo(1400,800);

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<% if oitem.FResultCount>0 then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<form name=frm2 method=post action="do_simpleiteminfoedit.asp">
	<input type=hidden name=menupos value="<%= menupos %>">
	<input type=hidden name=itemid value="<%= itemid %>">
	<input type=hidden name=itemoptionarr value="">
	<input type=hidden name=optisusingarr value="">
    <input type=hidden name=optdanjongynarr value="">
	<input type=hidden name=optremainnoarr value="">
	<input type=hidden name="optrackcodearr" value="">
    <input type=hidden name="suboptrackcodearr" value="">
	<input type="hidden" name="deliverytype" value="<%=oitem.FOneItem.Fdeliverytype%>">
	<input type="hidden" name="chkSRC" value="<%=returnValue%>">
	<tr>
	<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 border="0" class=a bgcolor=#BABABA>
			<tr height="25">

		<td width="120" bgcolor="#DDDDFF">상품명</td>
		<td colspan="2" bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">브랜드ID/브랜드명</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= oitem.FOneItem.Fmakerid %>/<%= oitem.FOneItem.FBrandName %>
			&nbsp;&nbsp;
			브랜드랙코드 : <%= prtidx %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">소비자가/매입가</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			&nbsp;&nbsp;
			<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %> %
			<% end if %>
		</td>
	</tr>

	<% if (oitem.FOneItem.FSailYn="Y") then %>
		<tr height="25">
			<td bgcolor="#DDDDFF">할인가/매입가</td>
			<td colspan="2" bgcolor="#FFFFFF">
				<font color="red">
					<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
					&nbsp;&nbsp;
					<% if (oitem.FOneItem.Forgprice<>0) then %>
				        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %>%
				    <% end if %>
				    할인
				</font>
				&nbsp;&nbsp;
				<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
				&nbsp;
				<% if oitem.FOneItem.FSellcash<>0 then %>
					<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
				<% end if %>
			</td>
		</tr>
	<% end if %>

	<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>
		<tr height="25">
			<td bgcolor="#DDDDFF">쿠폰가/매입가</td>
			<td colspan="2" bgcolor="#FFFFFF">
				<font color="green">
					<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>
					&nbsp;&nbsp;
					<%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
				</font>
			</td>
		</tr>
	<% end if %>

	<tr height="25">
		<td bgcolor="#DDDDFF">상품랙코드</td>
		<td bgcolor="#FFFFFF" width="270">
			<input type="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="8" maxlength="8" > (4 or 8자리 Fix)
		</td>
		<td rowspan="5" align="right" bgcolor="#FFFFFF">
			<img src="<%= oitem.FOneItem.FListImage %>" width="100" align="right">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">보조랙코드</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="subitemrackcode" value="<%= oitem.FOneItem.Fsubitemrackcode %>" size="8" maxlength="8" > (4 or 8자리 Fix)
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">사용옵션</td>
		<td bgcolor="#FFFFFF">
		(<%= oitem.FOneItem.FOptionCnt %> 개)
		&nbsp;
		<input type=button class="button" value="옵션추가/수정" onclick="popoptionEdit('<%= itemid %>');">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">배송구분</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsUpcheBeasong then %>
		<b>업체</b>배송
		<% else %>
		텐바이텐배송
		<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">상품 품절여부</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsSoldOut then %>
		<font color=red><b>품절</b></font>
		<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">평균 배송소요일</td>
		<td bgcolor="#FFFFFF" colspan="2">

			<% if (oitem.FOneItem.FavgDLvDate>-1) then %>
			    D+<%= oitem.FOneItem.FavgDLvDate+1 %>
			<% else %>
			    데이터 없음
			<% end if %>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:popItemAvgDlvGraph('<%= itemid %>');">[월별그래프]</a>&nbsp;
			<a href="javascript:popItemAvgDlvList('<%= itemid %>');">[상세리스트]</a>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">판매 시작일</td>
		<td bgcolor="#FFFFFF" colspan="2">
			<%= oitem.FOneItem.FsellSTDate %>
		</td>
	</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">상품 판매여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FSellYn="Y" then %>
					<input type="radio" name="sellyn" value="Y" checked >판매함
					<input type="radio" name="sellyn" value="S" >일시품절
					<input type="radio" name="sellyn" value="N" >판매안함
					<% elseif oitem.FOneItem.FSellYn="S" then %>
					<input type="radio" name="sellyn" value="Y" >판매함
					<input type="radio" name="sellyn" value="S" checked ><font color="red">일시품절</font>
					<input type="radio" name="sellyn" value="N" >판매안함
					<% else %>
					<input type="radio" name="sellyn" value="Y" >판매함
					<input type="radio" name="sellyn" value="S" >일시품절
					<input type="radio" name="sellyn" value="N" checked ><font color="red">판매안함</font>
					<% end if %>

                    <input type="button" class="button" value="히스토리" onClick="jsPopItemHistory(<%= itemid %>)">
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">상품 사용여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FIsUsing="Y" then %>
					<input type="radio" name="isusing" value="Y" checked >사용함
					<input type="radio" name="isusing" value="N" >사용안함
					<% else %>
					<input type="radio" name="isusing" value="Y" >사용함
					<input type="radio" name="isusing" value="N" checked ><font color="red">사용안함</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">제휴 사용여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FIsExtUsing="Y" then %>
					<input type="radio" name="isExtusing" value="Y" checked >사용함
					<input type="radio" name="isExtusing" value="N" >사용안함
					<% else %>
					<input type="radio" name="isExtusing" value="Y" >사용함
					<input type="radio" name="isExtusing" value="N" checked ><font color="red">사용안함</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">상품 단종여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.Fdanjongyn="Y" then %>
						<input type="radio" name="danjongyn" value="N" >생산중
						<input type="radio" name="danjongyn" value="S" >재고부족
						<input type="radio" name="danjongyn" value="Y" checked ><font color="red">단종품절</font>
						<input type="radio" name="danjongyn" value="M" >MD품절
					<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
						<input type="radio" name="danjongyn" value="N" >생산중
						<input type="radio" name="danjongyn" value="S" checked ><font color="red">재고부족</font>
						<input type="radio" name="danjongyn" value="Y" >단종품절
						<input type="radio" name="danjongyn" value="M" >MD품절
					<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
						<input type="radio" name="danjongyn" value="N" >생산중
						<input type="radio" name="danjongyn" value="S" >재고부족
						<input type="radio" name="danjongyn" value="Y" >단종품절
						<input type="radio" name="danjongyn" value="M" checked ><font color="red">MD품절</font>
					<% else %>
						<input type="radio" name="danjongyn" value="N" checked >생산중
						<input type="radio" name="danjongyn" value="S" >재고부족
						<input type="radio" name="danjongyn" value="Y" >단종품절
						<input type="radio" name="danjongyn" value="M" >MD품절
					<% end if %>
					<font color="#AAAAAA">
					<br> (상품판매에는 영향없슴 - 추가 입고예정 없을시 단종설정)
				</font>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">한정판매여부</td>
				<td bgcolor="#FFFFFF">
				<% if oitem.FOneItem.FLimitYn="Y" then %>
				<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">한정판매</font>
				<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">비한정판매
				(<%= oitem.FOneItem.FLimitNo %>-<%= oitem.FOneItem.FLimitSold %>=<%= oitem.FOneItem.FLimitNo-oitem.FOneItem.FLimitSold %>)
				<% else %>
				<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">한정판매
				<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">비한정판매
				<% end if %>
				<div id="dvDisp" style="display:<% if oitem.FOneItem.FLimitYn<>"Y" then %>none<%END IF%>;" >&nbsp;-> 한정노출여부:
					<input type="radio" name="limitdispyn" value="Y" <%IF oitem.FOneItem.Flimitdispyn="Y" or oitem.FOneItem.Flimitdispyn ="" THEN%>checked<%END IF%>>노출
					<input type="radio" name="limitdispyn" value="N" <%IF oitem.FOneItem.Flimitdispyn="N" THEN%>checked<%END IF%>>비노출</div>
				</td>
			</tr>
			</table>

		</td>
	</tr>
	<tr>
	    <td colspan="2" bgcolor="#FFFFFF">한정비교재고가 10미만일 경우는 재고파악 후 수기로 입력하시기 바랍니다.</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25" align="center" bgcolor="#FFDDDD" >
				<td width="50">옵션<br />코드</td>
				<td>옵션명</td>
				<td width="70">옵션<br />사용<br />여부</td>
                <td width="70">옵션<br />단종<br />구분</td>
				<td width="40">현재<br>한정</td>
				<td width="80">한정판매수량<br><input name="recalcuLimit" type="button" class="button" value="한정재계산" onclick="resetLimit();" <%= chkIIF(oitem.FOneItem.FLimitYn="N","disabled","") %>></td>
				<td width="40"><a href="javascript:TnPopItemStock('<%= itemid %>','');">한정<br />비교<br />재고</a></td>
				<td width="180">랙번호 / 보조랙</td>
			</tr>
			<% if oitemoption.FResultCount>0 then %>
				<% for i=0 to oitemoption.FResultCount - 1 %>
					<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
					<tr align="center" bgcolor="#EEEEEE">
					<% else %>
					<tr align="center" bgcolor="#FFFFFF">
					<% end if %>
						<td><%= oitemoption.FITemList(i).FItemOption %></td>
						<td><%= oitemoption.FITemList(i).FOptionName %></td>
						<td>
							<% if oitemoption.FITemList(i).FOptIsUsing="Y" then %>
							<input type="radio" id="optisusing0_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >Y <input type="radio" id="optisusing1_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >N
							<% else %>
							<input type="radio" id="optisusing0_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >Y <input type="radio" id="optisusing1_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">N</font>
							<% end if %>
						</td>
                        <td>
                            <select class="select" id="optdanjongyn_<%= i %>" name="optdanjongyn<%= oitemoption.FITemList(i).FItemOption %>" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn<>"N", "style='background-color:#FFBBDD'", "") %>>
                                <option value="N" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="N", "selected", "") %>>생산중</option>
                                <option value="S" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="S", "selected", "") %>>재고부족</option>
                                <option value="Y" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="Y", "selected", "") %>>옵션단종</option>
                                <option value="M" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="M", "selected", "") %>>MD품절</option>
                            </select>
                        </td>
						<td><%= oitemoption.FITemList(i).GetOptLimitEa %></td>
						<td>
							<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" dumistock="<%= oitemoption.FITemList(i).GetLimitStockNo %>" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitemoption.FITemList(i).GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','<%= oitemoption.FITemList(i).FItemOption %>');"><%= oitemoption.FITemList(i).GetLimitStockNo %></a></td>
						<td>
							<input type="text" id="optrackcode<%= oitemoption.FITemList(i).FItemOption %>" name="optrackcode" value="<%= oitemoption.FITemList(i).Foptrackcode %>" size="8" maxlength="8" >
                            <input type="text" id="suboptrackcode<%= oitemoption.FITemList(i).FItemOption %>" name="suboptrackcode" value="<%= oitemoption.FITemList(i).Fsuboptrackcode %>" size="8" maxlength="8" >
						</td>
					</tr>
				<% next %>
			<% else %>
					<tr align="center" bgcolor="#FFFFFF">
						<td>0000</td>
						<td colspan="3">옵션없음</td>
						<td><%= oitem.FOneItem.GetLimitEa %></td>
						<td>
							<input type="text" id="0000" dumistock="<%= oitem.FOneItem.GetLimitStockNo %>" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitem.FOneItem.GetLimitStockNo<10,"bgcolor='#6666EE'","") %> >
							<a href="javascript:TnPopItemStock('<%= itemid %>','');"><%= oitem.FOneItem.GetLimitStockNo %></a>
						</td>
						<td></td>
					</tr>
			<% end if %>
			</table>
		</td>
	</tr>

	<% IF oitem.FOneItem.Fsellyn = "N" THEN '판매안함 상태일때만 보여준다 %>
		<tr>
			<td   bgcolor="#FFFFFF">
				<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="0">
			 	<tr>
					<td>
						<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();" <%IF defaultcheck THEN%>checked<%END IF%>> 상품오픈예약:
						<input type="text" id="dSR" name="dSR" value="<%= FormatDateTime(oitem.FOneItem.Fsellreservedate,2) %>" size="12" class="input" />
						<select name="settime">
							<% for i=0 to 23 %>
							<option value="<%=Format00(2,i)%>"<% if Hour(oitem.FOneItem.Fsellreservedate)=i then response.write " selected" %>><%=Format00(2,i)%></option>
							<% next %>
						</select>시
						<img id="dSR_trigger" src="/images/admin_calendar.png" />
						  <div style="padding:3px">사용안함 상태일 경우 예약된 시간에 오픈이 되지 않습니다. <br>
					   텐바이텐 배송일 경우, 입고 확인 후 오픈예약이 가능합니다.  </div>
					   <script type="text/javascript">
						var CAL_SR = new Calendar({
							inputField : "dSR", trigger    : "dSR_trigger",
							onSelect: function() {
								this.hide();
							}
							, min: "<%=date()%>"
							, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					   </script>
					</td>
					</tR>
				</table>
			</td>
		</tr>
	<% END IF %>

	<tr align="center">
	    <td colspan="2" bgcolor="#FFFFFF">
			<input type="button" value="저장하기" onclick="SaveItem(frm2)" class="button">
			<input type="button" value=" 닫 기 " onclick="CloseWindow()" class="button">
		</td>
	</tr>
	<input type=hidden name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
	</form>
	</table>
<% else %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr bgcolor="#FFFFFF">
	    <td align="center">[검색 결과가 없습니다.]</td>
	</tr>
	</table>
<% end if %>

<Br>
<% if oitemrackoption.FTotalCount >0 then %>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oitemrackoption.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		※ 랙코드 변경 로그 / 검색결과 : <b><%= oitemrackoption.FTotalCount %></b>&nbsp;&nbsp; ※ 최대 50개까지 노출됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=80>수정</td>
	<td>변경내용</td>
	<td width=70>상품코드</td>
	<td width=60>옵션코드</td>
	<td>옵션명</td>
	<td width=80>예전<br>랙코드</td>
	<td width=80>예전<br>보조랙코드</td>
</tr>
<% for i=0 to oitemrackoption.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= left(oitemrackoption.FItemList(i).fregdate,10) %>
		<br><%= mid(oitemrackoption.FItemList(i).fregdate,12,22) %>
		<br><%= oitemrackoption.FItemList(i).fregadminid %>
	</td>
	<td align="left">
		<%= oitemrackoption.FItemList(i).fcomment %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fitemid %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fitemoption %>
	</td>
	<td align="left">
		<%= oitemrackoption.FItemList(i).Fitemoptionname %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).frackcodeByOption %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fsubRackcodeByOption %>
	</td>
</tr>   
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing
set oitemrackoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
