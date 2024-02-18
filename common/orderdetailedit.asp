<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim i
dim idx, orderserial
idx = requestCheckVar(request("idx"),10)

dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

if (ojumunDetail.Fresultcount<1) then
    ojumunDetail.FRectOldJumun = "on"
    ojumunDetail.SearchOneJumunDetail idx
end if

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if


if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

if ojumun.FResultCount<1 then
	response.write "해당되는 주문건이 없습니다."
	dbget.close() : response.end
end if

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
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

function EditDetail(detailidx,mode,comp) {
<% if ojumunDetail.FRectOldJumun = "on" or ojumun.FRectOldOrder = "on" then %>
    alert('과거내역 수정불가.');
    return;
<% end if %>
    var frm = document.frm;

	if (mode=="buycash") {
		if (!IsDigit(comp.value)) {
			alert("매입가는 숫자만 가능합니다.");
			comp.focus();
			return;
		}
	}else if (mode=="reducedPrice") {
		if (!IsDigit(comp.value)) {
			alert("쿠폰가는 숫자만 가능합니다.");
			comp.focus();
			return;
		}
	}else if (mode=="itemcost") {
		if (!IsDigit(comp.value)) {
			alert("쿠폰가는 숫자만 가능합니다.");
			comp.focus();
			return;
		}
    }else if (mode=="itemcostCouponNotApplied") {
		if (!IsDigit(comp.value)) {
			alert("판매가는 숫자만 가능합니다.");
			comp.focus();
			return;
		}
	}else if(mode=="isupchebeasong") {
	    if (frm.isupchebeasong.value=="Y") {
	        if (frm.omwdiv.value!="U") {
	            alert("매입구분과 배송구분이 일치하지 않습니다.");
	            return;
	        }
	    }else{
	        if (frm.omwdiv.value=="U") {
	            alert("매입구분과 배송구분이 일치하지 않습니다.");
	            return;
	        }
	    }

        if (frm.omwdiv.value=="U") {
            if ((frm.odlvType.value=="1")||(frm.odlvType.value=="4")) {
                alert("매입구분과 배송구분이 일치하지 않습니다.");
	            return;
            }
        }else{
            if ((frm.odlvType.value!="1")&&(frm.odlvType.value!="4")) {
                alert("매입구분과 배송구분이 일치하지 않습니다.");
	            return;
            }
        }

	}else if(mode=="songjangdiv") {
		// 잘못된 송장번호 삭제 및 EMS택배는 영문도 포함된 송장번호.
	    if (frm.songjangdiv.value.length<1) {
            alert("택배사를 선택하세요.\n\n송장번호가 없는 경우 기타로 입력하세요.");
			frm.songjangdiv.focus();
			return;
        }

        if (frm.songjangno.value == '') {
			alert("운송장번호를 입력하세요.\n\n송장번호가 없는 경우 기타로 입력하세요.");
			frm.songjangno.focus();
			return;
		}

		if (frm.applyallitem) {
			if (frm.applyallitem.checked == true) {
				if (confirm("해당 업체 [전체상품] 에 대해 송장번호를 입력합니다.\n\n진행하시겠습니까?") != true) {
					return;
				}
			}
		}

	}else if(mode=="currstate") {

	<% if ojumun.FOneItem.FIpkumdiv<4 then %>
	    alert("결제완료 이상만 상태변경가능.");
	    return;
	<% end if %>

		if (frm.currstate.value == '7') {
			// 잘못된 송장번호 삭제 및 EMS택배는 영문도 포함된 송장번호.
			if (frm.songjangdiv.value.length<1) {
				alert("택배사를 선택하세요.\n\n송장번호가 없는 경우 기타로 입력하세요.");
				frm.songjangdiv.focus();
				return;
			}

			if (frm.songjangno.value == '') {
				alert("운송장번호를 입력하세요.\n\n송장번호가 없는 경우 기타로 입력하세요.");
				frm.songjangno.focus();
				return;
			}

			/*
			if (frm.applyallitem) {
				if (frm.applyallitem.checked == true) {
					alert('한개의 상품에 대해서만 입력가능합니다.');
					frm.applyallitem.checked = false;
				}
			}
			*/
		}

	}else if (mode=="requiredetail") {

	}else if (mode=="itemno") {

	}else if (mode=="vatinclude") {
		if (comp.value == "") {
			alert("과세구분을 지정하세요.");
			return;
		}
    }else if (mode=="recalcmaster") {

	}else if (mode=="jungsan") {

    }else if (mode=="10x10logistics") {

	}else if (mode=="balju") {

    }else if (mode=="updmastercoupon") {

    }else if (mode=="ipkumdate") {

    }else if (mode=="additemid") {

	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm("수정 하시겠습니까?")) {
		frm.submit();
	}
}

function PopCurrentItemStock(itemid, itemoption) {
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?itemgubun=10&itemid=" + itemid + "&itemoption=" + itemoption,"PopCurrentItemStock","width=1000,height=600,resizable=yes,scrollbars=yes")
	popwin.focus();
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문아이템정보 수정</b>
		</td>
	</tr>

	<form name="frm" method="post" action="orderedit_process.asp">
	<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
	<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
	<input type="hidden" name="presongjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>">
	<input type="hidden" name="presongjangdiv" value="<%= ojumunDetail.FJumunDetail.Fsongjangdiv %>">
	<input type="hidden" name="mode" value="">
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">Idx.</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="120"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td>
			<%= ojumunDetail.FJumunDetail.Fitemid %>
			&nbsp;
			<input type="button" class="button" value="상품별재고현황" onClick="PopCurrentItemStock('<%= ojumunDetail.FJumunDetail.Fitemid %>', '<%= ojumunDetail.FJumunDetail.Fitemoption %>')">
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">옵션코드</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemoption %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">발주상태</td>
		<td>

		</td>
		<td>
		    <input type="button" class="button" value="결제완료전환" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'balju')">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
		<td>
            <input type="text" class="text" name="itemcostCouponNotApplied" value="<%= ojumunDetail.FJumunDetail.FitemcostCouponNotApplied %>" size="7" maxlength="9">
			<% if (ojumunDetail.FJumunDetail.Fitemcost<>ojumunDetail.FJumunDetail.FCurrsellcash) then %>
				(현판매가:<%= ojumunDetail.FJumunDetail.FCurrsellcash %>)
			<% end if %>
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="판매가수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemcostCouponNotApplied',frm.itemcostCouponNotApplied)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상품쿠폰가</td>
		<td>
            <input type="text" class="text" name="itemcost" value="<%= ojumunDetail.FJumunDetail.Fitemcost %>" size="7" maxlength="9">
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="쿠폰가수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemcost',frm.itemcost)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">보너스쿠폰가</td>
		<td>
            <input type="text" class="text" name="reducedPrice" value="<%= ojumunDetail.FJumunDetail.FreducedPrice %>" size="7" maxlength="9">
			<% if (ojumunDetail.FJumunDetail.Fitemcost<>ojumunDetail.FJumunDetail.FCurrsellcash) then %>
				(현판매가:<%= ojumunDetail.FJumunDetail.FCurrsellcash %>)
			<% end if %>
            * 할인액 텐바이텐부담
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="쿠폰가수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'reducedPrice',frm.reducedPrice)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">매입가</td>
		<td>
			<input type="text" class="text" name="buycash" value="<%= ojumunDetail.FJumunDetail.Fbuycash %>" size="7" maxlength="9">
			<% if ojumunDetail.FJumunDetail.Fitemcost<>0 then %>
			(<%= 100-CLng(ojumunDetail.FJumunDetail.Fbuycash/ojumunDetail.FJumunDetail.Fitemcost*10000/100) %> %)
			<% end if %>
			<% if (ojumunDetail.FJumunDetail.Fbuycash<>ojumunDetail.FJumunDetail.FCurrbuycash) then %>
				(현매입가:<%= ojumunDetail.FJumunDetail.FCurrbuycash %>)
			<% end if %>
            * 관리자 또는 회계팀 수정 가능
        </td>
		<td>
		    <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="매입가수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'buycash',frm.buycash)">
		    <% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
		<td>
			<select class="select" name="vatinclude" class="text">
        		<option value="">선택</option>
        		<option value="Y" <% if ojumunDetail.FJumunDetail.Fvatinclude="Y" then response.write "selected" %> >과세</option>
        		<option value="N" <% if ojumunDetail.FJumunDetail.Fvatinclude="N" then response.write "selected" %> >면세</option>
    		</select>
			* 정산내역이 있는 경우 수정불가!
    	</td>
		<td>
			<% if (C_ADMIN_AUTH) then %>
			<input type="button" class="button" value="과세구분수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'vatinclude',frm.vatinclude)">
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">마일리지</td>
		<td><%= ojumunDetail.FJumunDetail.Fmileage %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">구매수량</td>
		<td><input type="text" class="text" name="itemno" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9">개</td>
		<td>
		    <% if (C_ADMIN_AUTH) then %>
		    <input type="button" class="button" value="수량수정" <%= CHKIIF(ojumun.FOneItem.FIpkumdiv>6,"disabled","") %> onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemno',frm.itemno)">
		    <% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">배송구분</td>
		<td>
			<select class="select" name="isupchebeasong">
	    		<option value="Y" <% if ojumunDetail.FJumunDetail.Fisupchebeasong="Y" then response.write "selected" %> >업체배송
	    		<option value="N" <% if ojumunDetail.FJumunDetail.Fisupchebeasong="N" then response.write "selected" %> >자체배송
			</select>

	        <select class="select" name="omwdiv">
	            <option value="M" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="M","selected","") %> >매입
	            <option value="W" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="W","selected","") %> >위탁
	            <option value="U" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="U","selected","") %> >업체
	        </select>

	        <select class="select" name="odlvType">
	            <option value="1" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="1","selected","") %> >자체배송
	            <option value="2" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="2","selected","") %> >업체배송
	            <option value="4" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="4","selected","") %> >자체무료
	            <option value="5" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="5","selected","") %> >업체무료
	            <option value="7" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="7","selected","") %> >업체착불
	            <option value="9" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="9","selected","") %> >업체조건
	            <option value="6" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="6","selected","") %> >현장수령
	        </select>
		</td>
		<td>
		    <% if (C_ADMIN_AUTH) then %>
		    <input type="button" class="button" value="배송구분수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'isupchebeasong',frm.isupchebeasong)" >
		    <% end if %>
		</td>
	</tr>

	<% if ojumunDetail.FJumunDetail.Fisupchebeasong="Y" then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
		<td>
			<select class="select" name="currstate" class="text">
                <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
                <option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">출고이전 전환</option>
                <% else %>
        		<option value="" <% if ojumunDetail.FJumunDetail.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">미확인
        		<option value="2" <% if ojumunDetail.FJumunDetail.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">업체통보
        		<option value="3" <% if ojumunDetail.FJumunDetail.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">주문확인
        		<option value="7" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">출고완료
                <% end if %>
    		</select>
			* 정산내역이 있는 경우 수정불가!
    	</td>
		<td><input type="button" class="button" value="확인상태수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문통보일<br>(발주일)</td>
		<td><%= ojumun.FOneItem.FBaljuDate %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">업체확인일</td>
		<td><input type="text" class="text" name="upcheconfirmdate" value="<%= ojumunDetail.FJumunDetail.Fupcheconfirmdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">출고일</td>
		<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FJumunDetail.Fbeasongdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">택배정보</td>
		<td>
			<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FJumunDetail.Fsongjangdiv %>
			<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>" size="17" maxlength="20">
			<br /><input type="checkbox" name="applyallitem" value="Y"> 전상품적용
		</td>
		<td><input type="button" class="button" value="택배정보수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
	</tr>

	<% else %>

	<% if (C_ADMIN_AUTH or session("ssBctId") = "hasora" or session("ssBctId") = "boyishP" or session("ssBctId") = "oesesang52" or session("ssBctId") = "rabbit1693") then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
		<td>
			<select class="select" name="currstate" class="text">
                <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
                <option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">출고이전 전환</option>
                <% else %>
				<option value="" <% if ojumunDetail.FJumunDetail.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">미확인</option>
        		<option value="2" <% if ojumunDetail.FJumunDetail.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">물류통보</option>
        		<option value="3" <% if ojumunDetail.FJumunDetail.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">주문확인</option>
        		<option value="7" onChange="CheckConfirmDate(this);">출고완료</option>
                <% end if %>
    		</select>
			* skyer9,tozzinet,hasora,boyishP,oesesang52, rabbit1693 only
			<br>* 추가로 재고보정 필요
			<br>* 정산내역이 있는 경우 수정불가!
    	</td>
		<td><input type="button" class="button" value="확인상태수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)"></td>
	</tr>
    <% elseif (C_ADMIN_AUTH or C_CSPowerUser) then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
		<td>
            <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
            <select class="select" name="currstate" class="text">
        		<option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">출고이전 전환</option>
    		</select>
            <% else %>
            출고상품에만 사용가능
            <input type="hidden" name="currstate" value="<%= ojumunDetail.FJumunDetail.Fcurrstate %>">
            <% end if %>
			* CS관리자 권한
			<br>* 정산내역이 있는 경우 수정불가!
    	</td>
		<td>
            <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
            <input type="button" class="button" value="확인상태수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)">
            <% end if %>
        </td>
	</tr>
	<% else %>
	<input type="hidden" name="currstate" value="<%= ojumunDetail.FJumunDetail.Fcurrstate %>">
	<% end if %>

	<tr bgcolor="#FFFFFF" >
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문통보일<br>(발주일)</td>
		<td><%= ojumun.FOneItem.FBaljuDate %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">출고일</td>
		<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FJumunDetail.Fbeasongdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">택배정보</td>
		<td>
			<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FJumunDetail.Fsongjangdiv %>
			<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>" size="20" maxlength="20">
			<br /><input type="checkbox" name="applyallitem" value="Y"> 전상품적용
		</td>
		<td><input type="button" class="button" value="택배정보수정" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
	</tr>

	<% end if %>

	<% if ojumunDetail.FJumunDetail.Foitemdiv="06" then %>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">주문제작문구</td>
		<td>
		    <% if ojumunDetail.FJumunDetail.FItemNo=1 then %>
		    <textarea readonly class="textarea" name="requiredetail" cols="40" rows="3"><%= ojumunDetail.FJumunDetail.Frequiredetail %></textarea>
		    <% else %>
		    <% for i=0 to ojumunDetail.FJumunDetail.FItemNo-1 %>
		    <textarea readonly class="textarea" name="requiredetail<%=i%>" cols="40" rows="3"><%= splitValue(ojumunDetail.FJumunDetail.Frequiredetail,CAddDetailSpliter,i) %></textarea>
		    <% next %>
		    <% end if %>
		</td>
		<td>
		    <input type="button" class="button" value="주문제작문구수정" onclick="EditRequireDetail('<%= orderserial %>','<%= ojumunDetail.FJumunDetail.Fdetailidx %>')">
		</td>
	</tr>
	<% end if %>

    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">마스터쿠폰액</td>
		<td>
            마스터쿠폰액 오류
		</td>
		<td>
		    <input type="button" class="button" value="수정" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'updmastercoupon',null)">
		</td>
	</tr>

    <% if C_ADMIN_AUTH then %>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">주문마스터</td>
		<td>
		</td>
		<td>
		    <input type="button" class="button" value="재계산[관리자]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'recalcmaster',null)">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">정산</td>
		<td>
		</td>
		<td>
		    <input type="button" class="button" value="체크[관리자]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'jungsan',null)">
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">배송비</td>
		<td>
            10x10logistics 배송비추가
		</td>
		<td>
		    <input type="button" class="button" value="추가[관리자]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'10x10logistics',null)">
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">입금일</td>
		<td>
            <input type="text" class="text" name="ipkumdate" value="<% ''CHKIIF(ojumun.FOneItem.Fipkumdate="", "", FormatDate(ojumun.FOneItem.Fipkumdate, "0000.00.00 00:00:00")) %>" size="21" maxlength="30">
		</td>
		<td>
		    <input type="button" class="button" value="수정[관리자]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'ipkumdate',null)" disabled>
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">상품추가</td>
		<td>
            원주문 : <input type="text" class="text" name="orgorderserial" value="" size="10" maxlength="30">
            <br />
            디테일 : <input type="text" class="text" name="orgdetailidx" value="" size="10" maxlength="30">
		</td>
		<td>
		    <input type="button" class="button" value="수정[관리자]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'additemid',null)">
		</td>
	</tr>
    <% end if %>
	</form>
</table>


<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
