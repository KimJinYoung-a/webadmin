<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 초기 생성자 모름
'			 2010.05.25 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim ContractType , detailKey , i ,mode, sqlStr, ContractContents, ContractName ,subtype
	ContractType     = request("ContractType")
	mode             = request("mode")
	ContractContents = request("ContractContents")
	ContractName     = request("ContractName")
	detailKey        = request("detailKey")
	subtype          = request("subtype")

if (mode="editProtoType") then
    sqlStr = "update db_partner.dbo.tbl_partner_contractType"   & VbCrlf
    sqlStr = sqlStr & " set ContractContents='" & html2db(ContractContents) & "'" & VbCrlf
    sqlStr = sqlStr & " , ContractName='"&html2db(ContractName)&"'" & VbCrlf
    sqlStr = sqlStr & " , subtype='" & subtype & "'" & VbCrlf
    sqlStr = sqlStr & " where ContractType=" & ContractType

    dbget.Execute sqlStr
elseif (mode="regProtoType") then
    sqlStr = "insert into db_partner.dbo.tbl_partner_contractType"   & VbCrlf
    sqlStr = sqlStr & " (ContractName,ContractContents,subtype)" & VbCrlf
    sqlStr = sqlStr & " values("
    sqlStr = sqlStr & " '" & html2db(ContractName) & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & html2db(ContractContents) & "'" & VbCrlf
    sqlStr = sqlStr & " ,"&subtype& VbCrlf
    sqlStr = sqlStr & " )"

	'response.write sqlStr
    dbget.Execute sqlStr

elseif (mode="delKey") then
    sqlStr = "delete from db_partner.dbo.tbl_partner_contractDetailType"
    sqlStr = sqlStr & " where ContractType=" & ContractType
    sqlStr = sqlStr & " and detailKey='" & detailKey & "'"

    dbget.Execute sqlStr
end if

'//상단 리스트
dim ocontractProtoType
set ocontractProtoType = new CPartnerContract
ocontractProtoType.FPageSize = 40
ocontractProtoType.getValidContractProtoTypeList

'//변수 타입 리스트
dim ocontractDetailProtoType
set ocontractDetailProtoType = new CPartnerContract
ocontractDetailProtoType.FRectContractType = ContractType
ocontractDetailProtoType.getContractDetailProtoType

'//계약서 세부 내용
dim onecontractProtoType
set onecontractProtoType = new CPartnerContract
onecontractProtoType.FRectContractType = ContractType
onecontractProtoType.getOneContractProtoType
%>

<script language='javascript'>

function preViewContractProtoType(ContractType){
    var popwin = window.open('preViewCtrProtoType.asp?ContractType=' + ContractType,'preViewProtoType','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function DocDownloadContractProtoType(ContractType){
    var popwin = window.open('DocDownloadpreViewCtrProtoType.asp?ContractType=' + ContractType,'DocDownloadpreViewProtoType','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function pdfDownloadContractProtoType(ContractType){
    var popwin = window.open('pdfDownloadpreViewCtrProtoType.asp?ContractType=' + ContractType,'DocDownloadpreViewProtoType','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function checkNSubmit(frm){
    if (frm.ContractName.value.length<1){
        alert('계약서 명을 입력하세요.');
        frm.ContractName.focus();
        return;
    }
    if (frm.subtype.value.length<1){
        alert('계약서 구분을 선택 하세요.');
        frm.subtype.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

function regDetailProtoType(ContractType){
    var popwin = window.open('regDetailProtoType.asp?ContractType=' + ContractType,'regDetailProtoType','width=500,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function editDetailProtoType(ContractType,detailKey){
    var popwin = window.open('regDetailProtoType.asp?ContractType=' + ContractType + '&detailKey=' + detailKey,'regDetailProtoType','width=500,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function DelThis(ContractType,detailKey){
    var frm = document.frmSvr;

    if (confirm('삭제 하시겠습니까?')){
        frm.mode.value="delKey";
        frm.detailKey.value = detailKey;
        frm.submit();
    }
}

function RegNewProtoType(){
    location.href="?";
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="javascript:RegNewProtoType();"><img src="/images/icon_new_registration.gif" width="75" height="20" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
    <td>ID</td>
    <td>계약서<br>명칭</td>
    <td>계약서<br>구분</td>
    <td>등록일</td>
    <td>미리보기</td>
</tr>
<% for i=0 to ocontractProtoType.FResultCount -1 %>
<% if ContractType=CStr(ocontractProtoType.FItemList(i).FContractType) then %>
<tr bgcolor="#CCCCCC" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
    <td><%= ocontractProtoType.FItemList(i).FContractType %></td>
    <td><a href="?ContractType=<%= ocontractProtoType.FItemList(i).FContractType %>"><%= ocontractProtoType.FItemList(i).FcontractName %></a></td>
    <td><%= ocontractProtoType.FItemList(i).getSubTypeName %></td>
    <td><%= ocontractProtoType.FItemList(i).FRegDate %></td>
    <td>
        <a href="javascript:preViewContractProtoType('<%= ocontractProtoType.FItemList(i).FContractType %>');"><img src="/images/iexplorer.gif" width="21" border="0"></a>
        <!--&nbsp;<a href="javascript:DocDownloadContractProtoType('<%= ocontractProtoType.FItemList(i).FContractType %>');"><img src="/images/btn_word.gif" width="70" border="0"></a> -->
        <!--&nbsp;<a href="javascript:pdfDownloadContractProtoType('<%= ocontractProtoType.FItemList(i).FContractType %>');"><img src="/images/pdficon.gif" width="21" border="0"></a></td> -->


</tr>
<% next %>
</table>
<br>

<%
'/수정인 경우
if (onecontractProtoType.FResultCount>0) then
%>

	<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
	<form name="frmSvr" method="post" action="">
	<input type="hidden" name="ContractType" value="<%= onecontractProtoType.FOneItem.FContractType %>">
	<input type="hidden" name="mode" value="editProtoType">
	<input type="hidden" name="detailKey" value="">

	<tr bgcolor="#DDDDFF">
	    <td colspan="3">계약서 내용 수정</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="200">계약서 명</td>
	    <td colspan="2">
	        <input type="text" name="contractName" value="<%= onecontractProtoType.FOneItem.FcontractName %>">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="200">계약서 구분</td>
	    <td colspan="2">
	        <% drawSubTypeGubun "subtype" , onecontractProtoType.FOneItem.Fsubtype %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="3"><textarea name="ContractContents" cols="100" rows="10"><%= onecontractProtoType.FOneItem.FContractContents %></textarea></td>
	</tr>
	<tr bgcolor="#DDDDFF">
	    <td width="200">변수Type(KEY)</td>
	    <td align="right" colspan="2">
	    <!--
	    	<input type="button" value="변수추가" onClick="regDetailProtoType('<%= onecontractProtoType.FOneItem.FContractType %>');" class="button">
	    -->
	    </td>
	</tr>
	<% for i=0 to ocontractDetailProtoType.FresultCount-1 %>
	<tr bgcolor="#FFFFFF">
	    <td width="200">
	    	<a href="javascript:editDetailProtoType('<%= onecontractProtoType.FOneItem.FContractType %>','<%= ocontractDetailProtoType.FItemList(i).FdetailKey %>');">
	    	<%= ocontractDetailProtoType.FItemList(i).FdetailKey %></a>
	    </td>
	    <td ><%= ocontractDetailProtoType.FItemList(i).FdetailDesc %></td>
	    <td width="50">
	    	<a href="javascript:DelThis('<%= onecontractProtoType.FOneItem.FContractType %>','<%= ocontractDetailProtoType.FItemList(i).FdetailKey %>');">
	    	<img src="/images/icon_delete.gif" width="45" border="0"></a>
	    </td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td align="center" colspan="3"><input type="button" value="수정하기" onClick="frmSvr.submit();" class="button"></td>
	</tr>
	</form>
	</table>

<%
'//신규등록
elseif ContractType="" then
%>

	<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
	<form name="frmSvr" method="post" action="">
	<input type="hidden" name="mode" value="regProtoType">

	<tr bgcolor="#DDDDFF">
	    <td colspan="3">계약서 내용 신규등록</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="100">계약서 명 </td>
	    <td colspan="2">
	        <input type="text" name="ContractName" size="40" maxlength="40" value="" >
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="100">계약서 구분</td>
	    <td colspan="2">
	        <% drawSubTypeGubun "subtype" , subtype %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="3">
	        <textarea name="ContractContents" cols="100" rows="10"></textarea>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
	    <td align="center" colspan="3"><input type="button" value="신규저장" onClick="checkNSubmit(document.frmSvr);"  class="button"></td>
	</tr>
	</form>
	</table>

<% end if %>

<%
set ocontractProtoType = Nothing
set onecontractProtoType = Nothing
set ocontractDetailProtoType = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->