<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/catemanageCls.asp" -->
<%
Dim olist
Dim catecode
Dim page, i, Depth1Code, Depth1Name
Dim yyyy, mm, viewGubun
Dim cateCDArr

catecode	= requestCheckVar(Request("catecode"),16)
page 		= requestCheckVar(Request("page"),2)
yyyy		= requestCheckVar(Request("yyyy"),4)
mm			= requestCheckVar(Request("mm"),2)
viewGubun	= requestCheckVar(Request("gubun"),3)

If yyyy = "" Then yyyy = LEFT(date(), 4)
If mm = "" Then mm = Mid(date(),6,2)
If page = "" Then page = 1

SET olist = new CMDCategory
	olist.FPageSize		= 500
	olist.FCurrPage		= 1
	olist.FRectCatecode	= catecode
	olist.FRectyyyy		= yyyy
	olist.FRectmm		= mm
	olist.FRectGubun	= viewGubun
	olist.getMDPurpose1DepthList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function form_check(f){
	f.submit();
}
function saveCateUserid(){
	var frm;
	var sValue, tarMoney, proMoney, targetTF, proTF;
	frm = document.sfrm;
	sValue = "";
	tarMoney = "";
	proMoney = "";
	depth1TargetM = "";
	depth1ProfitM = "";
	targetTF = false;
	proTF = false

	// 목표액
	for (var i=0;i<frm.TarMoney.length;i++){
		if (tarMoney==""){
			if(frm.TarMoney[i].value == ''){
				frm.TarMoney[i].value = 0;
			}
			tarMoney = frm.TarMoney[i].value;
			depth1TargetM = frm.TarMoney[0].value;
		}else{
			if(frm.TarMoney[i].value == ''){
				frm.TarMoney[i].value = 0;
			}
			if(parseInt(depth1TargetM, 10) < parseInt(frm.TarMoney[i].value, 10)){
				targetTF = true;
			}
			tarMoney = tarMoney+","+frm.TarMoney[i].value;
		}
	}

	// 수익액
	for (var j=0;j<frm.ProMoney.length;j++){
		if (proMoney==""){
			if(frm.ProMoney[j].value == ''){
				frm.ProMoney[j].value = 0;
			}
			proMoney = frm.ProMoney[j].value;
			depth1ProfitM = frm.ProMoney[0].value;
		}else{
			if(frm.ProMoney[j].value == ''){
				frm.ProMoney[j].value = 0;
			}
			if(parseInt(depth1ProfitM, 10) < parseInt(frm.ProMoney[j].value, 10)){
				proTF = true;
			}
			proMoney = proMoney+","+frm.ProMoney[j].value;
		}
	}
	if(targetTF == true || proTF == true){
		if(confirm("1뎁스의 목표액이나 수익액이 2뎁스의 금액보다 작습니다.\n저장 하시겠습니까?")) {
			document.frmMoney.tarMoneyarr.value = tarMoney;
			document.frmMoney.proMoneyarr.value = proMoney;
			document.frmMoney.submit();
		}		
	}else{
		if(confirm("저장 하시겠습니까?")) {
			document.frmMoney.tarMoneyarr.value = tarMoney;
			document.frmMoney.proMoneyarr.value = proMoney;
			document.frmMoney.submit();
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="8" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="sfrm" method="POST">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<strong><%= yyyy %>년 <%= mm %>월 목표매출 관리</strong>
	</td>
</tr>
<% If olist.FTotalCount>0 then %>
<tr height="30" bgcolor="#FFFFFF" align="center">
	<td>Depth1합계</td>
	<td></td>
	<td></td>
	<td><%= FormatNumber(olist.FOneItem.FTotaltargetMoney,0)&"원" %></td>
	<td></td>
</tr>
<% End If %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td>카테고리코드</td>
    <td>카테고리1Depth</td>
    <td>카테고리2Depth</td>
    <td>목표액</td>
    <td>수익액</td>
</tr>
<%
If olist.FResultCount > 0 Then
	For i = 0 to olist.FResultCount -1
		If olist.FItemList(i).FDepth = 1 Then 
			Depth1Code = olist.FItemList(i).FCatecode
			Depth1Name = olist.FItemList(i).FCatename
		End If
		
		If Depth1Code = "" Then
			Depth1Code = LEFT(olist.FItemList(i).FCatecode,3)
			Depth1Name = fnGet1DepthCode(left(olist.FItemList(i).FCatecode,3))
		End If
%>
<tr align="center" <%= Chkiif(olist.FItemList(i).FDepth="1","bgcolor=SKYBLUE","bgcolor=FFFFFF") %>  height="25">
	<td align="left"><%= olist.FItemList(i).FCatecode %></td>
    <td align="left"><%= Depth1Name %></td>
    <td align="left">
	<%
		If CStr(Depth1Code) = CStr(Left(olist.FItemList(i).FCatecode,3)) Then
			If olist.FItemList(i).FDepth <> 1 Then
				response.write olist.FItemList(i).FCatename
			End If
		End If
	%>
    </td>
    <td>
		<input type="text" name="TarMoney" value="<%= Chkiif(olist.FItemList(i).FTargetMoney="", 0, olist.FItemList(i).FTargetMoney) %>" class="text" style="ime-mode:disabled" onkeypress="if ((event.keyCode<48) || (event.keyCode>57)) event.returnValue=false;">
    </td>
    <td>
		<input type="text" name="ProMoney" value="<%= Chkiif(olist.FItemList(i).FProfitMoney="", 0, olist.FItemList(i).FProfitMoney) %>" class="text" style="ime-mode:disabled" onkeypress="if ((event.keyCode<48) || (event.keyCode>57)) event.returnValue=false;">
    </td>
</tr>
<%
		cateCDArr = cateCDArr & olist.FItemList(i).FCatecode & ","
	Next
%>
<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="5"><input type="button" value="저장" class="button_s" onclick="saveCateUserid();"></td>
</tr>
<%
Else
%>
<tr align="center" height="50" bgcolor="FFFFFF">
	<td colspan="5">데이터가 없습니다.</td>
</tr>
<% End If %>
</form>
</table>
<% 
If Right(cateCDArr,1) = "," Then
	cateCDArr = Left(cateCDArr, Len(cateCDArr) - 1)
End If
%>
<form name="frmMoney" method="post" action="/admin/mdMenu/purpose/purpose_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="yyyy" value="<%= yyyy %>">
<input type="hidden" name="mm" value="<%= mm %>">
<input type="hidden" name="gubun" value="<%= viewGubun %>">
<input type="hidden" name="cateCDArr" value="<%= cateCDArr %>">
<input type="hidden" name="tarMoneyarr">
<input type="hidden" name="proMoneyarr">
<input type="hidden" name="catecode" value="<%= catecode %>">
</form>
<% SET olist = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->