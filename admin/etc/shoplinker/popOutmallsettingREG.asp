<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim mode, makerid, partnerid, padminId, strSQL, defaultFreeBeasongLimit, defaultDeliverPay, mall_name, mname
Dim pid, padmId, dFreeBeasongLimit, dDeliverPay
mode 	= request("mode")
makerid = request("makerid")

pid		= request("pid")
padmId	= request("padmId")
dFreeBeasongLimit	= request("dFreeBeasongLimit")
dDeliverPay			= request("dDeliverPay")
mall_name				= request("mall_name")
strSQL = ""
strSQL = strSQL & " SELECT A.partnerid, A.padminId, isnull(C.defaultFreeBeasongLimit,'') as defaultFreeBeasongLimit, isnull(C.defaultDeliverPay,'') as defaultDeliverPay "
strSQL = strSQL & " FROM db_user.dbo.tbl_user_c as C"
strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo as A on C.userid = A.partnerid "
strSQL = strSQL & " WHERE partnerid = '"&makerid&"'"
rsget.open strSQL, dbget, 1
If not rsget.EOF Then
	partnerid = rsget("partnerid")
	padminId = rsget("padminId")
	defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
	defaultDeliverPay = rsget("defaultDeliverPay")
End If
rsget.close

If mode = "S" Then
	If partnerid = "" Then
		response.write "<script language='JavaScript'>alert('브랜드 정보를 수정해주세요');location.replace('popOutmallsettingREG.asp')</script>"
	End If
ElseIf mode = "I" Then
	strSQL = ""
	strSQL = strSQL & " IF NOT Exists(SELECT makerid, mall_user_id FROM db_item.dbo.tbl_Shoplinker_OutmallControl where makerid = '"&pid&"') "
	strSQL = strSQL & " BEGIN "
	strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_Shoplinker_OutmallControl "
    strSQL = strSQL & " (makerid, mall_user_id, mall_name, defaultFreeBeasongLimit, defaultDeliverPay)"
    strSQL = strSQL & " VALUES ('"&pid&"', '"&trim(padmId)&"', '"&trim(mall_name)&"', '"&trim(dFreeBeasongLimit)&"', '"&trim(dDeliverPay)&"')"
    strSQL = strSQL & " END ELSE "
    strSQL = strSQL & " BEGIN "
	strSQL = strSQL & " UPDATE db_item.dbo.tbl_Shoplinker_OutmallControl SET "
    strSQL = strSQL & " mall_user_id = '"&trim(padmId)&"'"
    strSQL = strSQL & " ,mall_name = '"&trim(mall_name)&"'"
    strSQL = strSQL & " ,defaultFreeBeasongLimit = '"&trim(dFreeBeasongLimit)&"'"
    strSQL = strSQL & " ,defaultDeliverPay = '"&trim(dDeliverPay)&"'"
    strSQL = strSQL & " WHERE makerid = '"&pid&"' "
	strSQL = strSQL & " END "
    dbget.Execute strSQL
	response.write "<script language='JavaScript'>alert('저장되었습니다.');opener.location.reload();window.close();</script>"
ElseIf mode = "U" Then
	strSQL = ""
	strSQL = strSQL & " SELECT makerid, mall_user_id, mall_name, defaultFreeBeasongLimit, defaultDeliverPay "
	strSQL = strSQL & " FROM db_item.dbo.tbl_Shoplinker_OutmallControl"
	strSQL = strSQL & " WHERE makerid = '"&makerid&"'"
	rsget.open strSQL, dbget, 1
	If not rsget.EOF Then
		partnerid = rsget("makerid")
		padminId = rsget("mall_user_id")
		defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
		defaultDeliverPay = rsget("defaultDeliverPay")
		mname = rsget("mall_name")
	End If
	rsget.close
End If

If defaultFreeBeasongLimit = "" Then defaultFreeBeasongLimit = 0
If defaultDeliverPay = ""  Then defaultDeliverPay = 0
%>
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript">
function frmsearch(){
	if(document.frm.makerid.value == ""){
		document.frm.makerid.focus();
		alert('브랜드ID를 검색하세요');
		return false;
	}
	document.frm.mode.value="S";
	document.frm.submit();
}
function frmCheck(){
	if(document.frm2.pid.value == ""){
		alert('브랜드ID를 검색하세요');
		document.frm.makerid.focus();
		return false;
	}

	if(document.frm2.padmId.value == ""){
		document.frm2.padmId.focus();
		alert('제휴몰 ADMIN ID를 입력하세요');
		return false;
	}
	document.frm2.mode.value="I";
	document.frm2.submit();
}

function outSearchlID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }
    var popwin = window.open("/admin/member/popMeachulIDSearch.asp?pcuserdiv=999_50&usingonly=on&frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"jsoutID","width=800 height=400 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<br>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<% If mode <> "U" Then %>
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="mode">
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF">브랜드ID</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="makerid" value="<%=makerid%>">
		<input type="button" class="button" value="IDSearch" onclick="outSearchlID('frm','makerid')";>
	</td>
</tr>
<tr align="center">
	<td colspan="2" bgcolor="#FFFFFF"><input type="button" class="button" value="등록된 브랜드 검색" onclick="frmsearch();"></td>
</tr>
</table>
</form>
<% End If %>
<br>
<form name="frm2" method="GET" style="margin:0px;">
<input type="hidden" name="mode">
<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF" width="19%">브랜드ID</td>
	<td bgcolor="#FFFFFF"><%=partnerid%>
		<input type="hidden" name= "pid" value="<%=partnerid%>">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" width="19%">브랜드명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name= "mall_name" value="<%=mname%>"><br>
		<font color="RED">※반드시 샵링커 어드민의 기본정보관리 - 쇼핑몰SCM로그인의 쇼핑몰명을 입력하세요</font>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" width="19%">제휴몰ADMIN ID</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="padmId" value="<%=padminId%>"><br>
		<font color="RED">※반드시 샵링커 어드민의 기본정보관리 - 쇼핑몰SCM로그인의 로그인ID를 입력하세요</font>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" width="19%">기본 배송비 조건</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="dFreeBeasongLimit" value="<%=defaultFreeBeasongLimit%>" readonly>원 미만 구매시 배송료
		<input type="text" name="dDeliverPay" value="<%=defaultDeliverPay%>" readonly>원<br>
		<font color="BLUE">**제휴사(온라인) 제휴사 입점 정보**에 저장된 조건이 나옵니다. 해당 팝업에서 수정 불가능</font><br>
		샵링커 관리용이기 때문에 실제 배송비는 샵링커 어드민의 상품연동관리 -> 쇼핑몰 그룹정보 관리에서 제어함
	</td>
</tr>
<tr align="center">
	<td colspan="2" bgcolor="#FFFFFF"><input type="button" class="button" value="저장" onclick="frmCheck();"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
