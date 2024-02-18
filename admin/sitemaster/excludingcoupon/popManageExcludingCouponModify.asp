<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송비 부담금액 수정 팝업
' Hieditor : 2020.08.27 원승현 추가
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/excludingcoupon/excludingcouponcls.asp"-->
<%
Dim i, mode
Dim idx
dim oExcludingCouponView, loginUserId

idx = requestCheckvar(request("idx"), 50)

loginUserId = session("ssBctId")

if Trim(idx) = "" then
	response.write "<script>alert('정상적인 경로로 접근해주세요.');window.close();</script>"
	response.end
end If

'// halfdeliverypay View 데이터를 가져온다.
set oExcludingCouponView = new CgetExcludingCoupon
	oExcludingCouponView.FRectIdx = idx
	oExcludingCouponView.getExcludingCouponview()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
</head>
<body>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type='text/javascript'>
document.domain = "10x10.co.kr";

function frmedit(){
	var frm  = document.frm;

	if(confirm("수정하시겠습니까?")) {
		frm.submit();
	} else {
		return false;
	}
}

function checkLength(objname, maxlength)
{
	var objstr = objname.value;
	var objstrlen = objstr.length

	var maxlen = maxlength;
	var i = 0;
	var bytesize = 0;
	var strlen = 0;
	var onechar = "";
	var objstr2 = "";

	for (i = 0; i < objstrlen; i++)
	{
		onechar = objstr.charAt(i);

		if (escape(onechar).length > 4)
		{
			bytesize += 2;
		}
		else
		{
			bytesize++;
		}

		if (bytesize <= maxlen)
		{
			strlen = i + 1;
		}
	}

	if (bytesize > maxlen)
	{
		alert("허용된 문자열을 초과하였습니다.\n한글 기준 최대 "+maxlength/2+"자 까지 작성할 수 있습니다.");
		objstr2 = objstr.substr(0, strlen);
		objname.value = objstr2;
	}
	objname.focus();
}

function jsAddItemData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/common/pop_singleItemSelect.asp?target=frm&ptype=excludingcoupon','popAddItem','width=1000,height=600');
	winAddItem.focus();
}

function jsAddBrandData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/admin/member/popBrandSearch.asp?frmName=frm&compName=makerid&isjsdomain=o','popAddBrand','width=1000,height=600');
	winAddItem.focus();
}
</script>
<%' 팝업 사이즈 : 750*800 %>
<form name="frm" method="post" action="excludingCoupon_proc.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="idx" value="<%=oExcludingCouponView.FOneExcludingCoupon.Fidx%>">
<input type="hidden" name="excludingCouponType" value="<%=oExcludingCouponView.FOneExcludingCoupon.Ftype%>">
	<div class="popWinV17">
		<h1>수정</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>번호(idx) <strong class="cRd1"></strong></div></th>
					<td><%=oExcludingCouponView.FOneExcludingCoupon.Fidx%></td>
				</tr>
				<tr>
					<th><div>구분 <strong class="cRd1">*</strong></div></th>
					<td>
                        <%
                            If oExcludingCouponView.FOneExcludingCoupon.Ftype = "I" Then
                                response.write "상품"
                            ElseIf oExcludingCouponView.FOneExcludingCoupon.Ftype = "B" Then
                                response.write "브랜드"
                            End If
                        %>
					</td>
				</tr>
                <% If oExcludingCouponView.FOneExcludingCoupon.Ftype = "I" Then %>
                    <tr id="itemAddArea">
                        <th><div>상품등록 <strong class="cRd1">*</strong></div></th>
                        <td>
                            <input type="text" id="itemid" name="itemid" size="10"  value="<%=oExcludingCouponView.FOneExcludingCoupon.FItemID%>" />
                            <input type="button" value="상품검색" onclick="jsAddItemData();" style="width:100px;" />
                        </td>
                    </tr>
                <% ElseIf oExcludingCouponView.FOneExcludingCoupon.Ftype = "B" Then %>
                    <tr id="brandAddArea">
                        <th><div>브랜드 등록 <strong class="cRd1">*</strong></div></th>
                        <td>
                            <input type="text" id="makerid" name="makerid" size="10" value="<%=oExcludingCouponView.FOneExcludingCoupon.Fbrandid%>" />
                            <input type="button" value="브랜드 검색" onclick="jsAddBrandData();" style="width:100px;" />
                        </td>
                    </tr>
                <% End If %>
				<tr>
					<th><div>사용여부 <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" <% If oExcludingCouponView.FOneExcludingCoupon.Fisusing="N" Then %> checked <% End If %> /> 사용안함</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" <% If oExcludingCouponView.FOneExcludingCoupon.Fisusing="Y" Then %> checked <% End If %> /> 사용함</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>등록정보</div></th>
					<td>
						<span class="tPad05 col2"><%=oExcludingCouponView.FOneExcludingCoupon.Fadminid%>(<%=fnGetMyname(oExcludingCouponView.FOneExcludingCoupon.Fadminid)%>)<br/><%=oExcludingCouponView.FOneExcludingCoupon.Fregdate%></span>
					</td>
				</tr>
				<% If oExcludingCouponView.FOneExcludingCoupon.Flastadminid <> "" Then %>
				<tr>
					<th><div>최종수정</div></th>
					<td>
						<span class="tPad05 col2 cRd1"><%=oExcludingCouponView.FOneExcludingCoupon.Flastadminid%>(<%=fnGetMyname(oExcludingCouponView.FOneExcludingCoupon.Flastadminid)%>)<br/><%=oExcludingCouponView.FOneExcludingCoupon.Flastupdate%></span>
					</td>
				</tr>
				<% End If %>
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="수정" onclick="frmedit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
	set oExcludingCouponView = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
