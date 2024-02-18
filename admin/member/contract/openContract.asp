<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 계약 관리
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
dim groupid : groupid= requestCheckvar(request("groupid"),10)

dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=200
	ocontract.FCurrPage = 1
	ocontract.FRectGroupID = groupid
	'ocontract.FRectMakerid = makerid
	'ocontract.FRectContractno  = contractNo
	'ocontract.FRectContractState = ContractState
	ocontract.GetNewContractList

dim ogroupInfo
SET ogroupInfo = new CPartnerGroup
ogroupInfo.FRectGroupid = groupid
if (groupid<>"") then
    ogroupInfo.GetOneGroupInfo

    if (ogroupInfo.FResultCount<1) then
        response.write "해당 업체그룹 정보가 없습니다. "&groupid
        dbget.close(): response.end
    end if
end if

''마진 검토 리스트
dim oAddContractList
set oAddContractList = new CPartnerContract
oAddContractList.FPageSize=30
oAddContractList.FCurrPage = 1
oAddContractList.FRectGroupID = groupid
if (groupid<>"") then
    oAddContractList.GetCurrAddContractListCheckMargin
end if


dim i, noMatCnt, dsbleCnt
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="contract.js?v=1.00"></script>
<script language='javascript'>
$(document).on('change','input[name="chkAll"]',function() {
    $('.idChk').prop("checked" , this.checked);
});

function checkALL(comp){

}


function sendContract(inoMatCnt, dsbleCnt){
<% if session("ssbctId")<>"icommang" then %>
//alert('현재 오픈할 수 없습니다.');
//return;
<% end if %>
    var frm=document.frmOpen;

    var chkCnt = $('.idChk:checked').length;

    if (frm.dsbleCnt.value*1>0){
        alert('계약 불가한 조건이 있습니다.\n\n회색 줄 참조');
        return;
    }

    if (chkCnt<1){
        alert('선택된 계약서가 없습니다.');
        frm.chkAll.focus();
        return;
    }

    if ((frm.noMatCnt.value*1>0)&&(!confirm('현재 SCM 설정 마진과 일치하지 않는 내역이 있습니다.\n\n계속 하시겠습니까?\n\n계속 하시는경우 계약마진으로 현재 SCM마진이 변경됩니다.'))){
        return;
    }

 var signtype="";
   for(var i=0;i<frm.chkCtr.length;i++){
   	 	if (frm.chkCtr[i].checked){
	 		if (signtype ==""){
	   	 		signtype = frm.hidst[i].value;
	   	}else{
	   	 	 if (frm.hidst[i].value != signtype){
	   	 	 	alert("계약서 타입이 동일한 계약서만 일괄 발송 가능합니다.전자와 수기 또는 DocuSign 계약서를 따로 발송해주세요");
	   	 	 	return;
	   	 	 }
	     	}
	   	  frm.signtype.value = signtype;
  	 	}    	
   }

    <%'' DocuSign 일 경우 ctropen이 아닌 ctropendocusign 으로 ctrReg_Process.asp 페이지에 보낸다. %>
    if (frm.signtype.value=="3") {
        frm.mode.value="ctropendocusign";
        $("#submitButton").attr('disabled',true);
    }

    if (confirm('선택 계약서를 발송(오픈)하시겠습니까?')){
        frm.submit();
    }

}

function jsEcSubmit(ecCtrSeq, companyno){
	document.frmecView.cont_seq.value = ecCtrSeq;
	document.frmecView.corp_id.value = companyno;
	document.frmecView.submit();
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<form name="frm" method="get" action="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">

    업체그룹코드 : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="10" >
    <input type="button" class="button" value="Code검색" onclick="popSearchGroupID(this.form.name,'groupid');" >
    &nbsp;&nbsp;
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<form name="frmecView" method="post" action="<%=FecUrl%>/w20/contractView.do" target="_blank"> 
	<input type="hidden" name="remote_id" value="<%=FecID%>" />  <!-- 작성자 LOGIN ID -->
	<input type="hidden" name="cont_seq" value="" />  <!-- 계약서 번호 -->
	<input type="hidden" name="corp_id" value="" /> <!-- 계약을 화인하려는 사업자번호 -->
 </form> 
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmOpen" method="post" action="ctrReg_Process.asp">
<input type="hidden" name="groupid" value="<%=groupid%>">
<input type="hidden" name="mode" value="ctropen">
<input type="hidden" name="reguserid" value="<%=session("ssBctID")%>">
<input type="hidden" name="signtype" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20" ><input type="checkbox" name="chkAll"></td>
    <td width="50" >계약서타입</td>
    <td width="110" >계약서 명</td>
    <td width="100" >계약서번호</td>
    <td width="80" >업체명</td>
    <td width="100" >브랜드ID</td>
    <td width="100" >판매처</td>
    <td width="100" >마진</td>
    <td width="70" >계약일</td>
    <td width="70" >상태</td>
    <td width="80" >등록자</td>
    <td width="70" >등록일</td>
    <td width="90" >비고</td>
</tr>
<% dim signtype: signtype = 0

 if ocontract.FResultCount>0 then %>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
    <td ><input type="checkbox" class="<%=CHKIIF(ocontract.FITemList(i).IsCtrOpenValidState,"idChk","")%>" name="chkCtr" value="<%= ocontract.FITemList(i).FctrKey %>" <%=CHKIIF(ocontract.FITemList(i).IsCtrOpenValidState,"","disabled")%> onClick="checkALL(this)"></td>
    <td><%if ocontract.FItemList(i).FecCtrSeq <> "" and not isNull(ocontract.FItemList(i).FecCtrSeq) and ocontract.FItemList(i).FecCtrSeq <> "0" then%>
    		전자(<%=ocontract.FItemList(i).FecCtrSeq %>)
    		<%signtype =2 %>
    		<%else%>
                <%If ocontract.FItemList(i).FsignType = "D" Then %>
                    DocuSign
                    <%signtype =3 %>
                <% Else %>
    		        수기
                    <%signtype =1 %>                    
                <% End If %>
    		<%end if%>
    	<input type="hidden" name="hidst" value="<%=signtype%>">
    </td>
    <td><%= ocontract.FITemList(i).FContractName %></td>
    <td align="center"><a href="javascript:modiContract('<%= ocontract.FITemList(i).FctrKey %>');"><%= CHKIIF(isNULL(ocontract.FITemList(i).FctrNo) or ocontract.FITemList(i).FctrNo="","-",ocontract.FITemList(i).FctrNo) %></a></td>
    <td><%= ocontract.FITemList(i).FcompanyName %></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorMarginStr %></td>
    <td align="center"><%= ocontract.FITemList(i).FcontractDate %></td>
    <td align="center"><font color="<%= ocontract.FITemList(i).GetContractStateColor %>"><%= ocontract.FITemList(i).GetContractStateName %></font></td>
    <td align="center"><%= ocontract.FITemList(i).FRegUserName %></td>
    <td ><%= LEFT(ocontract.FITemList(i).FregDate,10) %></td>
    <td align="center">
    	<%if  ocontract.FITemList(i).FecCtrSeq <>"" and not  isNull(ocontract.FITemList(i).FecCtrSeq ) and ocontract.FITemList(i).FecCtrSeq <> "0" then%>
        <img src="/images/documents_icon.png" style="cursor:pointer;" onClick="jsEcSubmit('<%=ocontract.FITemList(i).FecCtrSeq%>','<%= replace(ocontract.FItemList(i).FcompanyNo,"-","")%>');" >
         
        <%end if%>
        <% If ocontract.FITemList(i).FsignType = "D" Then %>
            <img src="/images/browser_icon.png" style="cursor:pointer;" onClick="dnWebAdmDocu('<%=ocontract.FITemList(i).FctrKey %>');">
        <% Else %>
            <img src="/images/browser_icon.png" style="cursor:pointer;" onClick="dnWebAdm('<%=ocontract.FITemList(i).FctrKey %>');">
        <% End If %>
        <img src="/images/pdf_icon.png" style="cursor:pointer;" onClick="dnPdfAdm('<%=ocontract.FITemList(i).getPdfDownLinkUrlAdm %>');">
    </td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">오픈할 계약서가 없습니다.</td>
</tr>
<% end if %>
</table>
<p>

<!-- 마진 검토 --> 
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td colspan="4">계약정보</td>
    <td colspan="4">SCM설정정보</td>
    <td colspan="4">상품정보</td>
    <td rowspan="2">비고</td>
</tr>
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td>브랜드ID</td>
    <td>판매처</td>
    <td>매입구분</td>
    <td>계약마진</td>


    <td>매입구분</td>
    <td>계약마진</td>

    <td>대표매입구분</td>
    <td>대표계약마진</td>

    <td>사용수</td>
    <td>마진</td>
    <td>판매수</td>
    <td>마진</td>

</tr> 
<%
noMatCnt=0
dsbleCnt=0
%>

<% for i=0 to oAddContractList.FresultCount-1 %>
<%
if (oAddContractList.FItemList(i).isreqCheckMargin) then
    noMatCnt=noMatCnt+1
end if

if (oAddContractList.FItemList(i).isDisabledMWMargin) then
    dsbleCnt=dsbleCnt+1
end if
%> 
<tr bgcolor="<%=CHKIIF(oAddContractList.FItemList(i).isDisabledMWMargin,"#CCCCCC","#FFFFFF")%>" align="center">

    <td bgcolor="#FFFFFF" ><%=oAddContractList.FItemList(i).FMakerid %></td>
    <td bgcolor="#FFFFFF" ><%=oAddContractList.FItemList(i).getSellplaceName %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).isreqCheckMW,"bgcolor='#DD7777'","")%> ><%=oAddContractList.FItemList(i).getContractMwDivStr %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).isreqCheckMargin,"bgcolor='#DD7777'","")%> ><%=oAddContractList.FItemList(i).getContractMarginStr %></td>

    <td><%=fnMaeipdivName(oAddContractList.FItemList(i).FMaeipdiv) %></td>
    <td><%=oAddContractList.FItemList(i).getSCMDefaultmargineStr %></td>

    <% if (oAddContractList.FItemList(i).Fsellplace="ON") then %>
    <td><% if (oAddContractList.FItemList(i).Fcontractmwdiv=oAddContractList.FItemList(i).FMjmaeipdiv) then %><%=fnMaeipdivName(oAddContractList.FItemList(i).FMjmaeipdiv) %><% end if %></td>
    <td><% if (oAddContractList.FItemList(i).Fcontractmwdiv=oAddContractList.FItemList(i).FMjmaeipdiv) then %><%=oAddContractList.FItemList(i).getMjContractMarginStr %><% end if %></td>
    <% else %>
    <td><%=fnMaeipdivName(oAddContractList.FItemList(i).FMjmaeipdiv) %></td>
    <td><%=oAddContractList.FItemList(i).getMjContractMarginStr %></td>
    <% end if %>

    <td><%=FormatNumber(oAddContractList.FItemList(i).FuseitemCnt,0) %></td>
    <td><%=CLNG(oAddContractList.FItemList(i).Fuseitemmargin*100)/100 %></td>
    <td><%=FormatNumber(oAddContractList.FItemList(i).FsellitemCnt,0) %></td>
    <td><%=CLNG(oAddContractList.FItemList(i).Fsellitemmargin*100)/100 %></td>
    <td>

    </td>
</tr> 
<% next %>
<input type="hidden" name="noMatCnt" value="<%=noMatCnt%>">
<input type="hidden" name="dsbleCnt" value="<%=dsbleCnt%>">
</table> 
<p>
<% if ocontract.FResultCount>0 then %>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#FFFFFF">
<tr bgcolor="#FFFFFF">
    <td align="center">
        <table width="80%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
        <tr bgcolor="#FFFFFF">
            <td width="100">계약담당자</td>
            <td><%=session("ssBctCName")%></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td>계약수신자</td>
            <td><%=ogroupInfo.FOneItem.Fmanager_name%></td>
            <td><%=ogroupInfo.FOneItem.Fmanager_phone%></td>
            <td><input type="checkbox" name="ckHp" value="on" checked >발송 <input type="text" name="mngHp" value="<%=ogroupInfo.FOneItem.Fmanager_hp%>" class="text" size="15"></td>
            <td><input type="checkbox" name="ckEmail" value="on" checked >발송 <input type="text" name="mngEmail" value="<%=ogroupInfo.FOneItem.Fmanager_email%>" class="text" size="22"></td>
        </tr>
        </table>
    </td>
</tr>
<input type="hidden" name="mngName" value="<%=ogroupInfo.FOneItem.Fmanager_name%>">
</table>
<% end if %>
<p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#FFFFFF">
<tr bgcolor="#FFFFFF">
    <td align="center">
    <% if ocontract.FResultCount>0 then %>
    <input type="button" class="button" value="선택 계약서 발송" id="submitButton" name="submitButton" onClick="sendContract()">
    &nbsp
    <input type="button" class="button" value="이메일 미리보기" onClick="preViewSendContract('<%=groupid%>','1')">
    &nbsp
    <input type="button" class="button" value="전자계약 이메일 미리보기" onClick="preViewSendContract('<%=groupid%>','2')">
    <% else %>
    <input type="button" class="button" value="닫기" onClick="window.close()">
    <% end if %>
    </td>
</tr>
</form>
</table>

<%
SET oAddContractList=Nothing
SET ogroupInfo=Nothing
SET ocontract=Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->