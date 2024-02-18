<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="EUC-KR" %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->
<%
dim idx, mode, regtype, regdate
dim cMailzine, ArrTemplateInfo, ix, scriptTXT

idx = requestCheckVar(request("idx"),32)
mode = requestCheckVar(request("mode"), 32)
regtype = requestCheckVar(request("regtype"), 32)
regdate = requestCheckVar(request("regdate"), 32)

'템플릿 정보 가져오기
set cMailzine = new CMailzineList
cMailzine.FRectRegType = regtype
cMailzine.frectidx = idx
ArrTemplateInfo=cMailzine.fnMailzineTemplateContents
set cMailzine = nothing
scriptTXT=""
%>
<% If isArray(ArrTemplateInfo) Then %>
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
    <% For ix=0 To UBound(ArrTemplateInfo,2) %>
    <tr bgcolor="#FFFFFF" height="25">
        <td align="center" width="150"><%=ArrTemplateInfo(1, ix)%></td>
        <td>
            <% if ArrTemplateInfo(0,ix)="20" or ArrTemplateInfo(0,ix)="21" or ArrTemplateInfo(0,ix)="22" or ArrTemplateInfo(0,ix)="23" then %>
                <input type="button" class="button" value="이미지업로드" onClick="jsSetImg('<%=ArrTemplateInfo(3,ix)%>','img<%=ix+1%>');return false;"> <span id="img<%=ix+1%>"></span>
                <input type="hidden" name="img<%=ix+1%>" value="<%=ArrTemplateInfo(3,ix)%>">
                <% if ArrTemplateInfo(3,ix) <> "" then %>
                <input type="button" onclick="delimg('<%=ix+1%>');" class="button" value="이미지삭제">
                <% end if %>
                <% if ArrTemplateInfo(3,ix) <> "" then %>
                <textarea name="imagemap<%=ix+1%>" rows="10" class="textarea" style="width:95%;"><%=ArrTemplateInfo(2,ix)%></textarea>
                <% else %>
                <textarea name="imagemap<%=ix+1%>" rows="10" class="textarea" style="width:95%;"><map name="ImgMap<%=ix+1%>"></map></textarea>
                <% end if %>
            <% elseif ArrTemplateInfo(0,ix)="24" then %>
                <input type="text" name="evt_code<%=ix+1%>" class="input" size="10" value="<%=ArrTemplateInfo(2,ix)%>">
                <% scriptTXT = scriptTXT + "if(frm.evt_code"&ix+1&".value==""""){alert('주말특가 이벤트코드를 입력하세요.');frm.evt_code"&ix+1&".focus();return false;}else{if(frm.arrevtcode.value!=""""){frm.arrevtcode.value=frm.arrevtcode.value+','+frm.evt_code"&ix+1&".value;}else{frm.arrevtcode.value=frm.evt_code"&ix+1&".value;}return true;}" %>
            <% elseif ArrTemplateInfo(0,ix)="25" then %>
                <input type="text" name="evt_code<%=ix+1%>" class="input" size="10" value="<%=ArrTemplateInfo(2,ix)%>">
                <% scriptTXT = scriptTXT + "if(frm.evt_code"&ix+1&".value==""""){alert('메인 이벤트코드를 입력하세요.');frm.evt_code"&ix+1&".focus();return false;}else{if(frm.arrevtcode.value!=""""){frm.arrevtcode.value=frm.arrevtcode.value+','+frm.evt_code"&ix+1&".value;}else{frm.arrevtcode.value=frm.evt_code"&ix+1&".value;}return true;}" %>
            <% else %>
                <textarea class="textarea" cols="20" rows="6" name="contents<%=ix+1%>"><%=ArrTemplateInfo(2,ix)%></textarea>
            <% end if %>
            <input type="hidden" name="idx<%=ix+1%>" value="<%=ArrTemplateInfo(4,ix)%>">
        </td>
    </tr>
    <% next %>
</table>
<script>
function fnEvtCodeCheck(){
var frm = document.frm;
<% if scriptTXT="" then %>
    frm.arrevtcode.value="0";
    return true;
<% else %>
    frm.arrevtcode.value="";
    <%=scriptTXT%>
<% end if %>
}
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->