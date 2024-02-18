<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim id
id = requestcheckVar(request("id"),9)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if

dim divcd
if (ocsaslist.FResultcount>0) then
    divcd = ocsaslist.FOneItem.FDivCd
end if


''환불정보
dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = ocsaslist.FOneItem.FId
orefund.GetOneRefundInfo

response.write "<br><br>시스템팀 문의 : 사용중지 페이지!!"
dbget.close()
response.end

%>

<script language='javascript'>
//업체 추가 정산 삭제시
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";

    frm.buf_totupchejungsandeliverypay.value = frm.buf_refunddeliverypay.value*1 + frm.add_upchejungsandeliverypay.value*1;

}


//추가 정산 관련
function conFirmSave(frm){
    if (frm.add_upchejungsandeliverypay){
        if (frm.add_upchejungsandeliverypay.value == ""){
            alert('추가정산배송비를 입력하세요.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*0 != 0){
            alert('숫자만 가능합니다.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='직접입력')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('추가 정산액이 있는경우 브랜드 아이디가 지정되어야 합니다. ');
                return;
            }

            //주문 내역에 아이디가 있는 경우만.

        }else{
            <% if (divcd="A700") then %>
            //alert('추가 정산액을 입력하세요.');
            //frm.add_upchejungsandeliverypay.focus();
            //return;
            <% end if %>
        }
    }

    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

//추가정산배송비
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//추가정산배송비 사유
function Change_add_upchejungsancause(comp){
    if (comp.value=="직접입력") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}
</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    <form name="frmaction" method="post" action="pop_cs_action_process.asp">
    <input type="hidden" name="mode" value="addupchejungsanEdit">
    <input type="hidden" name="id" value="<%= id %>">
	<tr bgcolor="FFFFFF">
	    <td colspan="2"><strong>* 업체 추가 정산 내역</strong></td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">브랜드ID</td>
	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    <% if (divcd="A700") then %>
	    <input type="button" class="button" value="브랜드ID검색" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    <% end if %>
	    </td>
	</tr>

	<tr bgcolor="FFFFFF">
	    <td width="100">회수배송비</td>
	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">추가정산배송비</td>
	    <td ><input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">원
	    &nbsp;<select name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
	    <option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>사유선택
	    <option value="추가배송비" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="추가배송비","selected","") %> >추가배송비
	    <option value="추가운임" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="추가운임","selected","") %>>추가운임
	    <option value="직접입력" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>직접입력
	    </select>

	    <span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'><input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" ></span>
	    <a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    </td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td width="100">총정산배송비</td>
	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
    <td align="center">
    <input type="button" value="저장" onClick="conFirmSave(frmaction);">
    </td>
</tr>
</table>
<%
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
