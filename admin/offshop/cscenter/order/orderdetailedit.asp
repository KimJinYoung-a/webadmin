<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim i ,detailidx, masteridx ,ojumunDetail
	detailidx = requestCheckVar(request("detailidx"),10)

set ojumunDetail = new COrder
	ojumunDetail.frectdetailidx = detailidx
	ojumunDetail.fSearchOneJumunDetail()

	if ojumunDetail.ftotalcount > 0 then
		masteridx = ojumunDetail.FOneItem.fmasteridx
	else
		response.write "<script language='javascript'>"
		response.write "	alert('주문마스터 테이블 값이 없습니다');"
		response.write "	self.close();"
		response.write "</script>"
		response.end
	end if

dim ojumun
set ojumun = new COrder
	if (masteridx <> "") then
	    ojumun.FRectmasteridx = masteridx
	    ojumun.fQuickSearchOrderMaster()
	end if

%>

<script language='javascript'>

window.resizeTo(600,600);

var oldConfirmDate = "";
var oldBeasongDate = "";

function CheckConfirmDate(comp){
    if (comp.value==""){
        oldConfirmDate = comp.form.upcheconfirmdate.value;
        oldBeasongDate = comp.form.beasongdate.value;
        comp.form.upcheconfirmdate.value = "";
    }else{
        if (oldConfirmDate!=""){
            comp.form.upcheconfirmdate.value = oldConfirmDate;
        }

        if (oldBeasongDate!=""){
            comp.form.beasongdate.value = oldBeasongDate;
        }
    }
}

//수정
function EditDetail(detailidx,mode,comp){
    var frm = document.frm;

	if(mode=="currstate"){
		<% if ojumun.FOneItem.FIpkumdiv<4 then %>
		    alert('결제완료 이상만 상태변경가능.');
		    return;
		<% end if %>
    }else if(mode=="songjangdiv"){
        if (frm.songjangdiv.value.length<1){
            alert('택배사를 선택하세요.');
			frm.songjangdiv.focus();
			return;
        }

        if (!IsDigit(frm.songjangno.value)){
			alert('운송장번호는 숫자는 가능합니다.');
			frm.songjangdiv.focus();
			return;
		}
	}else if (mode=="itemno"){

	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a">
<form name="frm" method="post" action="/admin/offshop/cscenter/order/order_process.asp">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FOneItem.Fdetailidx %>">
<input type="hidden" name="masteridx" value="<%= ojumunDetail.FOneItem.Fmasteridx %>">
<input type="hidden" name="mode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문아이템정보 수정</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">Idx.</td>
	<td><%= ojumunDetail.FOneItem.Fdetailidx %></td>
	<td width="120"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
	<td>
		<%= ojumunDetail.FOneItem.fitemgubun%>-<%=CHKIIF(ojumunDetail.FOneItem.fitemid>=1000000,Format00(8,ojumunDetail.FOneItem.fitemid),Format00(6,ojumunDetail.FOneItem.fitemid))%>-<%=ojumunDetail.FOneItem.fitemoption %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">옵션코드</td>
	<td><%= ojumunDetail.FOneItem.Fitemoption %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
	<td><%= ojumunDetail.FOneItem.Fitemname %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
	<td><%= ojumunDetail.FOneItem.Fitemoptionname %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
	<td><%= ojumunDetail.FOneItem.Fmakerid %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
	<td>
		<%= ojumunDetail.FOneItem.fsellprice %>
		<% if (ojumunDetail.FOneItem.fsellprice<>ojumunDetail.FOneItem.FCurrsellcash) then %>
			(현판매가:<%= ojumunDetail.FOneItem.FCurrsellcash %>)
		<% end if %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">구매수량</td>
	<td><input type="text" class="text" name="itemno" value="<%= ojumunDetail.FOneItem.Fitemno %>" size="3" maxlength="9">개</td>
	<td>
	    <% if (C_ADMIN_AUTH) then %>
	    	<input type="button" class="button" value="수량수정" <%= CHKIIF(ojumun.FOneItem.FIpkumdiv>6,"disabled","") %> onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'itemno',frm.itemno)">
	    <% end if %>
	</td>
</tr>
<% if ojumunDetail.FOneItem.Fisupchebeasong="Y" then %>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
	<td>
		<select class="select" name="currstate" class="text">
    		<option value="" <% if ojumunDetail.FOneItem.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">미확인
    		<option value="2" <% if ojumunDetail.FOneItem.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">업체통보
    		<option value="3" <% if ojumunDetail.FOneItem.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">주문확인
    		<option value="7" <% if ojumunDetail.FOneItem.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">출고완료
		</select>
	</td>
	<td>
		<input type="button" class="button" value="확인상태수정" onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'currstate',frm.currstate)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문통보일<br>(발주일)</td>
	<td><%= ojumun.FOneItem.FBaljuDate %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">업체확인일</td>
	<td><input type="text" class="text" name="upcheconfirmdate" value="<%= ojumunDetail.FOneItem.Fupcheconfirmdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">출고일</td>
	<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FOneItem.Fbeasongdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>

<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">택배정보</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FOneItem.Fsongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FOneItem.Fsongjangno %>" size="20" maxlength="20">
	</td>
	<td><input type="button" class="button" value="택배정보수정" onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문통보일<br>(발주일)</td>
	<td><%= ojumun.FOneItem.FBaljuDate %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">출고일</td>
	<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FOneItem.Fbeasongdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">택배정보</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FOneItem.Fsongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FOneItem.Fsongjangno %>" size="20" maxlength="20">
	</td>
	<td></td>
</tr>
<% end if %>
</form>
</table>

<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->