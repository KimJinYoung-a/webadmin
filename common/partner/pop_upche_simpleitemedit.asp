<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품정보
' Hieditor : 2009.04.07 서동석 생성
'			 2011.04.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/partner/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" --><!--html--> 
<%
dim itemid ,i
	itemid = requestCheckvar(request("itemid"),16)  ''requestCheckvar 2016/02/11

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

''2017/06/19 추가 by eastone  itemid=7171719721 형태?
If LEN(itemid)>9 Then
	response.write "<script>"
	response.write "	alert('잘못된 상품코드입니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

Dim vQuery, vIsOK
''vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
''rsget.open vQuery,dbget,1
''2015/06/18 서동석
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
set oitem = new CItemInfo

oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

''2016/02/11 추가.
if (oitem.FResultCount<1) then
    response.write "<script>"
	response.write "	alert('잘못된 상품코드이거나 해당상품이 없습니다.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if
%>

<script language='javascript'>

function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && ((e.name.substring(0,"optlimitno".length) == "optlimitno")||(e.name.substring(0,"optlimitsold".length) == "optlimitsold"))) {
				e.disabled = (comp.value=="N");
		  }
  	}

}

function SaveItem(frm){
	frm.itemoptionarr.value = ""
	frm.optlimitnoarr.value = ""
	frm.optlimitsoldarr.value = ""
	frm.optisusingarr.value = ""

    var option_isusing_count = 0;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optlimitno".length)) == "optlimitno"){

		  	    if (!IsDigit(e.value)){
		  	        alert('한정수량은 숫자만 가능합니다.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + e.id + "," ;
				frm.optlimitnoarr.value = frm.optlimitnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                }
		  	}

		  	if ((e.name.substring(0,"optlimitsold".length)) == "optlimitsold") {
		  	    if (!IsDigit(e.value)){
		  	        alert('한정수량은 숫자만 가능합니다.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.optlimitsoldarr.value = frm.optlimitsoldarr.value + e.value + "," ;
			}

			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		}
  	}
    if (option_isusing_count < 1) {
        alert("모든 옵션을 사용안함으로 할수 없습니다. 상품 판매여부를 판매안함으로 변경해주세요");
        //alert(frm.itemoptionarr.value);
        return;
    }

<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
    if (frm.reqstring.value == "") {
        alert("수정사유를 입력해주세요.");
        return;
    }
<% end if %>

<%
	If vIsOK = "x" Then
		If oitem.FOneItem.FSellYn <> "Y" Then
%>
			if(frm.sellyn[0].checked)
			{
				alert("상품고시내용이 모두 입력되어 있지 않은 상태입니다.\n모두 입력하셔야 판매함으로 수정 가능합니다.\n모두 입력하신 뒤 이 창을 새로 여시거나 새로고침 하시면 수정이 가능합니다.");
				return;
			}
<%
		End If
	End If
%>

	var ret = confirm('저장 하시겠습니까?');

	if(ret){
		frm.submit();
	}
}

function PopOptionEdit(itemid){
	var popwin = window.open('/common/partner/pop_adminitemoptionedit.asp?itemid=' + itemid,'PopOptionEdit','width=800 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function editItemInfo(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/partner/itemmaster/item_infomodify.asp?' + param ,'editItemInfo','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : 타이틀 영역하단에 searchWrap이 올 경우엔 bgNone 클래스 삭제 -->
			<h2>상품수정</h2>
			<ul class="txtList">
				<li>텐배(텐바이텐배송)상품은 텐바이텐 확인후 <span class="cRd1">다음날 수정사항이 반영</span>됩니다.</li>
				<li>업배(업체배송) 상품인 경우 <span class="cRd1">즉시반영</span>됩니다.</li>
				<li>가격이나, 상품명 등 기타 수정 하실 내용은 <span class="cRd1">담당엠디</span>에게 문의해 주세요.</li>
			</ul>  
			</div>
		<div class="cont">  
			<h3>옵션/한정/판매관련</h3>
			<% if oitem.FResultCount>0 then %> 
			<form name="frm2" method="post" action="/common/partner/do_upche_simpleiteminfoedit.asp">
			<input type="hidden" name="itemid" value="<%= itemid %>">
			<input type="hidden" name="itemoptionarr" value="">
			<input type="hidden" name="optisusingarr" value="">
			<input type="hidden" name="optlimitnoarr" value="">
			<input type="hidden" name="optlimitsoldarr" value="">
			<table class="tbType1 writeTb tMar10">
					<colgroup>
						<col width="15%" /><col width="" />
					</colgroup>
					<tbody> 
					<tr> 
						<th><div>상품코드</div></td>
						<td><%= itemid %></td>
					</tr>
					<tr>
						<th><div>상품명</div></td>
						<td><%= oitem.FOneItem.Fitemname %></td>	
					</tr>
					<tr>
						<th><div>브랜드</div></td>
						<td><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
					</tr>
					<tr>
						<th><div>판매가/매입가</div></td>
						<td>
						<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
						</td>
					</tr>
					<tr>
						<th><div>매입구분</div></td>
						<td>
						<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
						&nbsp;
						<% if oitem.FOneItem.FSellcash<>0 then %>
						<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>사용옵션</div></td>
						<td>
						(<%= oitem.FOneItem.FOptionCnt %> 개)
						&nbsp;
						<% if oitem.FOneItem.IsUpcheBeasong then %>
						<input type=button value="옵션수정" onclick="PopOptionEdit('<%= itemid %>');" class="btn3 btnIntb">
						<% else %>
						<span class="cRd1">- 옵션 추가/삭제는 담당MD</span>에게 문의하세요.
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>배송구분</div></td>
						<td>
						<% if oitem.FOneItem.IsUpcheBeasong then %>
						<b>업체</b>배송
						<% else %>
						텐바이텐배송
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>상품 품절여부</div></td>

						<td>
						<% if (oitem.FOneItem.IsSoldOut) or (oitem.FOneItem.FSellYn="S") then %>
						<span class="cRd1"><strong>품절</strong></span>
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>상품 판매여부</div></td>
						<td>
						<% if oitem.FOneItem.FSellYn="Y" then %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" checked >판매함
						<input type="radio" name="sellyn" value="S" class="formRadio" >일시품절
						<input type="radio" name="sellyn" value="N" class="formRadio" >판매안함
						<% elseif oitem.FOneItem.FSellYn="S" then %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" >판매함
						<input type="radio" name="sellyn" value="S" class="formRadio" checked ><font color="blue">일시품절</font>
						<input type="radio" name="sellyn" value="N" class="formRadio" >판매안함
						<% else %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" >판매함
						<input type="radio" name="sellyn" value="S" class="formRadio" >일시품절
						<input type="radio" name="sellyn" value="N" class="formRadio" checked ><font color="red">판매안함</font>
						<% end if %>
						<% If vIsOK = "x" Then %>
					    	&nbsp;&nbsp;<input type="button" class=btn3 btnIntb" value="상품고시내용입력" style="width:110px;" onClick="editItemInfo('<%=itemid%>');">
						<% End If %>
						</td>
					</tr>
					<input type="hidden" name="isusing" value="<%= oitem.FOneItem.FIsUsing %>">
					<tr>
						<th><div>한정판매여부</div></td>
						<td>
						<% if oitem.FOneItem.FLimitYn="Y" then %>
						<input type="radio" class="formRadio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">한정판매</font>
						<input type="radio" class="formRadio" name="limityn" value="N" onclick="EnabledCheck(this)">비한정판매
						<% else %>
						<input type="radio" class="formRadio" name="limityn" value="Y" onclick="EnabledCheck(this)">한정판매
						<input type="radio" class="formRadio" name="limityn" value="N" checked onclick="EnabledCheck(this)">비한정판매
						<% end if %>
						</td>
					</tr>
				</table>
				<table class="tbType1 listTb">
					<thead>
						<tr> 
								<th><div>옵션명</div></th>
								<th><div>옵션사용여부</div></th>
								<th><div>한정수량 - 판매수량 = 한정재고</div></th>
								<th><div>비고</div></th>
						</tr>
					</thead>
					<tbody>
							<% if oitemoption.FResultCount>0 then %>
								<% for i=0 to oitemoption.FResultCount - 1 %>
									<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
									<tr bgcolor="#EEEEEE">
									<% else %>
									<tr bgcolor="#FFFFFF">
									<% end if %>
										<td><%= oitemoption.FITemList(i).FOptionName %>(<%= oitemoption.FITemList(i).FItemOption %>)</td>
										<td>
											<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
											<input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >사용함 <input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >사용안함
											<% else %>
											<input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >사용함 <input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><span class="cRd1">사용안함</span>
											<% end if %>
										</td>
										<td>
										<input type="text" class="formTxt" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
										-
										<input type="text" class="formTxt" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitsold<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
										=
										<input type="text" class="formTxt" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 disabled >
									</td>
									<td>
									<% if (oitemoption.FITemList(i).FOptIsUsing="N") or (oitemoption.FITemList(i).Foptsellyn="N") or (oitemoption.FITemList(i).Foptlimityn="Y" and oitemoption.FITemList(i).GetOptLimitEa<1) then %>
									<span class="cRd1">품절</span>
									<% end if %>
									</td>
									</tr>
								<% next %>
							<% else %>
								<tr>
									<td colspan="2">옵션없음 (0000)</td>
									<td>
									<input type="text" class="formTxt" id="0000" name="optlimitno" value="<%= oitem.FOneItem.FLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
									-
									<input type="text" class="formTxt" id="0000" name="optlimitsold" value="<%= oitem.FOneItem.FLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
									=
									<input type="text" class="formTxt" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 disabled >
								</td>
								<td>
								    <% if oitem.FOneItem.isSoldOut() then %>
								    <span class="cRd1">품절</span>
								    <% end if %>
								</td>
								</tr>
							<% end if %>
						</tbody>
					</table>
			 <input type="hidden" name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
			 <table class="tbType1 writeTb tMar10">
					<colgroup>
						<col width="15%" /><col width="" />
					</colgroup>
					<tbody> 
					<tr> 
						<th><div>이미지</div></th>
						<td>
							<img src="<%= oitem.FOneItem.FListImage %>" width=100>
						</td>
					</tr>
			<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
			<tr>
				<th><div>수정사유</div></th>
				<td>
				  <input type="text" class="formTxt" name="reqstring" value="" size="30">
				  <p class="tMar05 fs11">(ex: 절판, 재고일시부족(입고예정일 2003-05-15), 재입고..)</p>
				</td>
			</tr>
			<% end if %>
			</form>
		</table>
		<div class="tPad15 ct"> 
				<% if (oitem.FOneItem.Fmwdiv = "U") then %>
					<input type="button" value=" 저장하기 " onclick="SaveItem(frm2)" class="btn3 btnRd" />
				<% else %>
					<input type="button" value=" 수정요청 " onclick="SaveItem(frm2)" class="btn3 btnRd" />
				<% end if %>	
			</div>   
<% end if %>
		</div>
	</div>
</div>
</body>
</html>
<%
set oitemoption = Nothing
set oitem = Nothing
%>
 
<!-- #include virtual="/lib/db/dbclose.asp" -->