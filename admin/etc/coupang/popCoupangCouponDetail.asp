<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim midx, itemid, couponId
midx 		= request("midx")
couponId	= request("couponId")
itemid  	= request("itemid")
If NOT isNumeric(midx) Then
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');window.close();</script>"
	dbget.close()	:	response.End
End If

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function delCateItem(v)
{
	$("#delIdx").val(v);
	document.frm.target = "xLink";
	document.frm.submit();
}

function popCateSelect(){
	$.ajax({
		url: "/admin/etc/coupang/act_CategorySelect.asp",

		cache: false,
		success: function(message) {
			$("#lyrCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
function frm2Submit(f){
	if(confirm("저장 하시겠습니까?")) {
		f.target = "xLink";
		f.submit();
	}
}
function CoupangCouponDetailProcess(){
	if(confirm("아래 저장된 내용으로 전송 하시겠습니까?")) {
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "COUPONDETAILREG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
	}
}
function CoupangCouponDeleteDetailProcess(){
	if(confirm("아래 저장된 내용으로 전송 하시겠습니까?")) {
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "COUPONDETAILDEL";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
	}
}
</script>

<h2>쿠폰 옵션 등록</h2>
<input type="button" class="button" value="등록API" onclick="CoupangCouponDetailProcess();" />
<p>
<form name="frm" method="post" action="procCoupangCoupon.asp" onsubmit="return false;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" id="mode" name="mode" value="cateDetail">
<input type="hidden" id="delIdx" name="delIdx" value="">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td><%= getCategory(midx) %></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popCateSelect();"></td>
		</tr>
		</table>
		<div id="lyrCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
</table>
</form>
<br />

<form name="frm2" method="post" action="procCoupangCoupon.asp" onsubmit="return false;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" id="mode" name="mode" value="ItemDetail">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품" style="cursor:help;">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td>
				<%= getItemTextArea(midx) %>
				<input type="button" value="저장" class="button" onclick="frm2Submit(this.form);"/>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br /><br />
<h2>쿠폰 옵션 삭제</h2>
<input type="button" class="button" value="삭제API" onclick="CoupangCouponDeleteDetailProcess();" />
<p>
<form name="frm2" method="post" action="procCoupangCoupon.asp" onsubmit="return false;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" id="mode" name="mode" value="ItemDeleteDetail">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품" style="cursor:help;">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td>
				<%= getItemDeleteTextArea(midx) %>
				<input type="button" value="저장" class="button" onclick="frm2Submit(this.form);"/>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>


<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cksel" value="<%= midx %>">
<input type="hidden" name="couponId" value="<%= couponId %>">
<input type="hidden" name="cmdparam" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
