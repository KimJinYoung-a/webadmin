<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim falg
falg = request("falg")

%>
<script language='javascript'>
function MrOk(){
    if (document.all.rejectmsg.value.length<1){
        alert('사유를 선택하세요.');
        document.all.rejectmsg.focus();
        return;
    }else{
        document.all.ret.value = document.all.rejectmsg.value;
    }
    
    if (document.all.rejectmsg.value=="직접입력"){
        if (document.all.rejectmsg_Text.value.length<1){
            alert('사유를 입력 하세요.');
            document.all.rejectmsg_Text.focus();
            return;
        }else{
            document.all.ret.value = document.all.rejectmsg_Text.value;
        }
    }
    window.close();
}

function MrCancel(){
    document.all.ret.value = '';
    window.close();
}

function ChgCombo(comp){
    if (comp.value=="직접입력"){
        document.all.divtext.style.display = "inline";
    }else{
        document.all.divtext.style.display = "none";
    }
}

</script>
<BODY bgcolor="#ffffff" OnUnload="window.returnValue = document.all.ret.value;">
<INPUT type="hidden" name="ret">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="a">
<% if (falg="1") then %>
<!-- 등록보류 -->
<tr height="30">
    <td align="center">등록 보류 사유 선택</td>
</tr>
<tr height="30">
    <td align="center">
        <select name="rejectmsg" onChange="ChgCombo(this);">
        <option value="">선택
        <option value="이미지 등록 불량">이미지 등록 불량
        <option value="상품 설명 부족">상품 설명 부족
        <option value="직접입력">----직접입력----
        </select>
    </td>
</tr>

<% elseif (falg="2") then %>
<!-- 등록 불가 사유 -->
<tr height="30">
    <td align="center">등록 불가 사유 선택</td>
</tr>
<tr height="30">
    <td align="center">
        <select name="rejectmsg" onChange="ChgCombo(this);">
        <option value="">선택
        <option value="동일상품 판매중">동일상품 판매중
        <option value="직접입력">----직접입력----
        </select>
    </td>
</tr>

<% end if %>
<tr height="30">
    <td id="divtext" style="display=none;" align="center">
        <input type="text" name="rejectmsg_Text" size="30" maxlength="100">
    </td>
</tr>

<tr height="30">
    <td align="center">
        <input type="button" class="button" value="확인" onclick="MrOk()">
        <input type="button" class="button" value="취소" onclick="MrCancel()">
    </td>
</tr>
</table>
</body>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
