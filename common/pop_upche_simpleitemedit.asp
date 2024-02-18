<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품정보
' Hieditor : 2009.04.07 서동석 생성
'			 2011.04.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%
dim itemid ,i
	itemid = requestCheckvar(request("itemid"),10)  ''requestCheckvar 2016/02/11

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
        alert("모든 옵션을 사용안함으로 할수 없습니다. 상품정보를 사용안함으로 변경하거나, 전시안함 변경하세요.");
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
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + itemid,'PopOptionEdit','width=700 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function editItemInfo(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/designer/itemmaster/upche_item_infomodify.asp?' + param ,'editItemInfo','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr>
	<td align="left">
		<strong>상품정보 수정</strong><br>
		<br>- 텐배(텐바이텐배송)상품은 텐바이텐 확인후 <font color=red>다음날 수정사항이 반영</font>됩니다.
		<br>- 업배(업체배송) 상품인 경우 <font color=red>즉시반영</font>됩니다.
		<br>- 가격이나, 상품명 등 기타 수정 하실 내용은 <font color=red>담당엠디</font>에게 문의해 주세요.
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="frm2" method="post" action="do_upche_simpleiteminfoedit.asp">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="optisusingarr" value="">
<input type="hidden" name="optlimitnoarr" value="">
<input type="hidden" name="optlimitsoldarr" value="">
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">상품코드</td>
	<td width=76% bgcolor="#FFFFFF"><%= itemid %></td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">상품명</td>
	<td width=76% bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">브랜드</td>
	<td bgcolor="#FFFFFF"><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">판매가/매입가</td>
	<td bgcolor="#FFFFFF">
	<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">매입구분</td>
	<td bgcolor="#FFFFFF">
	<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
	&nbsp;
	<% if oitem.FOneItem.FSellcash<>0 then %>
	<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">사용옵션</td>
	<td bgcolor="#FFFFFF">
	(<%= oitem.FOneItem.FOptionCnt %> 개)
	&nbsp;
	<% if oitem.FOneItem.IsUpcheBeasong then %>
	<input type=button value="옵션수정" onclick="PopOptionEdit('<%= itemid %>');" class="button">
	<% else %>
	<font color=red>* 옵션 추가/삭제는 담당MD</font>에게 문의하세요.
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">배송구분</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.IsUpcheBeasong then %>
	<b>업체</b>배송
	<% else %>
	텐바이텐배송
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">상품 품절여부</td>
	<td bgcolor="#FFFFFF">
	<% if (oitem.FOneItem.IsSoldOut) or (oitem.FOneItem.FSellYn="S") then %>
	<font color=red><b>품절</b></font>
	<% end if %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">상품 판매여부</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.FSellYn="Y" then %>
	<input type="radio" name="sellyn" value="Y" checked >판매함
	<input type="radio" name="sellyn" value="S" >일시품절
	<input type="radio" name="sellyn" value="N" >판매안함
	<% elseif oitem.FOneItem.FSellYn="S" then %>
	<input type="radio" name="sellyn" value="Y" >판매함
	<input type="radio" name="sellyn" value="S" checked ><font color="blue">일시품절</font>
	<input type="radio" name="sellyn" value="N" >판매안함
	<% else %>
	<input type="radio" name="sellyn" value="Y" >판매함
	<input type="radio" name="sellyn" value="S" >일시품절
	<input type="radio" name="sellyn" value="N" checked ><font color="red">판매안함</font>
	<% end if %>
	<% If vIsOK = "x" Then %>
    	&nbsp;&nbsp;<input type="button" class="button" value="상품고시내용입력" style="width:110px;" onClick="editItemInfo('<%=itemid%>');">
	<% End If %>
	</td>
</tr>
<input type="hidden" name="isusing" value="<%= oitem.FOneItem.FIsUsing %>">
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF">한정판매여부</td>
	<td bgcolor="#FFFFFF">
	<% if oitem.FOneItem.FLimitYn="Y" then %>
	<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">한정판매</font>
	<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">비한정판매
	<% else %>
	<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">한정판매
	<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">비한정판매
	<% end if %>
	</td>
</tr>
<tr>
	<td colspan="2" height="25" bgcolor="#FFFFFF">
		<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
		<tr bgcolor="#FFDDDD">
			<td height="25">옵션명</td>
			<td width="100">옵션사용여부</td>
			<td>한정수량 - 판매수량 = 한정재고</td>
			<td width="40">비고</td>
		</tr>
		<% if oitemoption.FResultCount>0 then %>
			<% for i=0 to oitemoption.FResultCount - 1 %>
				<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
				<tr bgcolor="#EEEEEE">
				<% else %>
				<tr bgcolor="#FFFFFF">
				<% end if %>
					<td height="25"><%= oitemoption.FITemList(i).FOptionName %>(<%= oitemoption.FITemList(i).FItemOption %>)</td>
					<td>
						<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
						<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >사용함 <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >사용안함
						<% else %>
						<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >사용함 <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">사용안함</font>
						<% end if %>
					</td>
					<td>
					<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
					-
					<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitsold<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
					=
					<input type="text" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 disabled >
				</td>
				<td>
				<% if (oitemoption.FITemList(i).FOptIsUsing="N") or (oitemoption.FITemList(i).Foptsellyn="N") or (oitemoption.FITemList(i).Foptlimityn="Y" and oitemoption.FITemList(i).GetOptLimitEa<1) then %>
				<font color=red>품절</font>
				<% end if %>
				</td>
				</tr>
			<% next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td height="25" colspan="2">옵션없음 (0000)</td>
				<td>
				<input type="text" id="0000" name="optlimitno" value="<%= oitem.FOneItem.FLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
				-
				<input type="text" id="0000" name="optlimitsold" value="<%= oitem.FOneItem.FLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
				=
				<input type="text" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 disabled >
			</td>
			<td>
			    <% if oitem.FOneItem.isSoldOut() then %>
			    <font color=red>품절</font>
			    <% end if %>
			</td>
			</tr>
		<% end if %>
		</table>
	</td>
</tr>
<input type="hidden" name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
<tr>
	<td width=80 bgcolor="#DDDDFF">이미지</td>
	<td bgcolor="#FFFFFF">
	<img src="<%= oitem.FOneItem.FListImage %>" width=100>
	</td>
</tr>
<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
<tr>
	<td width=80 bgcolor="#DDDDFF">수정사유</td>
	<td bgcolor="#FFFFFF">
	  <input type="text" name="reqstring" value="" size="30"><br>(ex: 절판, 재고일시부족(입고예정일 2003-05-15), 재입고..)
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="#FFFFFF" colspan=2 align="center">
		<% if (oitem.FOneItem.Fmwdiv = "U") then %>
      		<input type="button" value="저장하기" onclick="SaveItem(frm2)" class="button">
		<% else %>
     		<input type="button" value="수정요청" onclick="SaveItem(frm2)" class="button">
		<% end if %>
		<input type="button" value=" 닫 기 " onclick="CloseWindow()" class="button">
	</td>
</tr>
</form>
</table>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->