<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->

<%
dim page
page=request("page")
if page="" then page=1

dim iting
set iting = new CTingItemList
iting.FPageSize = 100
iting.FCurrPage = page
iting.GetAllTingItemList

dim ix
%>
<script language="javascript">
function CheckNEditTing(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
				if (e.value.length<1){
					alert('필수 입력 사항입니다.');
					e.focus();
					return;
				}
			}

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
			if (!IsDigit(e.value)){
				alert('숫자만 가능합니다.');
				e.focus();
				return;
			}
		}
	}

	var ret = confirm('수정하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function checkNAddting(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
				if (e.value.length<1){
					alert('필수 입력 사항입니다.');
					e.focus();
					return;
				}
			}

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
			if (!IsDigit(e.value)){
				alert('숫자만 가능합니다.');
				e.focus();
				return;
			}
		}
	}
	var ret = confirm('추가하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>
<table width="960" border="1" cellpadding="0" cellspacing="0" class="a">
<tr >
	<td width="50" align="center">아이템ID</td>
	<td width="50" align="center">이미지</td>
	<td width="120" align="center">아이템명</td>
	<td width="70" align="center">팅포인트(정)</td>
	<td width="70" align="center">팅포인트(비)</td>
	<td width="70" align="center">구매구분</td>
	<td width="70" align="center">한정판매</td>
	<td width="64" align="center">총한정수량</td>
	<td width="64" align="center">판매수량</td>
	<td width="60" align="center">남은수량</td>
	<td width="70" align="center">전시(사용)여부</td>
	<td width="70" align="center">판매여부</td>
	<td width="70" align="center">이벤트구분</td>
	<td width="70" align="center">Evt_CPCode</td>
	<td width="50" align="center">수정</td>
</tr>
<% for ix=0 to iting.FResultCount-1 %>
<form name="frm_<%= iting.FTingList(ix).FID %>" method="post" action="dotingedit.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="id" value="<%= iting.FTingList(ix).FID %>">
<tr>
	<td align="center">
		<input type="text" name="itemid" value="<%= iting.FTingList(ix).FItemID %>" size="7" maxlength="7" readonly >
	</td>
	<td align="center"><img src=<%= iting.FTingList(ix).FImageSmall %>></td>
	<td align="center"><%= iting.FTingList(ix).FItemName %></td>
	<td align="center">
		<input type="text" name="tingpoint" value="<%= iting.FTingList(ix).FTingPoint %>" size=6>
	</td>
	<td align="center">
		<input type="text" name="tingpoint_b" value="<%= iting.FTingList(ix).FTingPoint_B %>" size=6>
	</td>
	<td align="center">
		<select name="userclass">
			<option value="A" <% if iting.FTingList(ix).FUserClass="A" then response.write "selected" %> >정팅,준팅</option>
			<option value="Y" <% if iting.FTingList(ix).FUserClass="Y" then response.write "selected" %> >정팅</option>
			<option value="N" <% if iting.FTingList(ix).FUserClass="N" then response.write "selected" %> >정팅,준팅,노팅</option>
		</select>
	</td>
	<td align="center">
		<select name="limitdiv">
			<option value="0" <% if iting.FTingList(ix).FLimitDiv="0" then response.write "selected" %> >비한정판매</option>
			<option value="1" <% if iting.FTingList(ix).FLimitDiv="1" then response.write "selected" %> >수량한정</option>
			<option value="2" <% if iting.FTingList(ix).FLimitDiv="2" then response.write "selected" %> >일별한정</option>
			<option value="3" <% if iting.FTingList(ix).FLimitDiv="3" then response.write "selected" %> >월별한정</option>
		</select>
	</td>
	<td align="center">
		<input type="text" name="limitea" value="<%= iting.FTingList(ix).FLimitea %>" size=6>
	</td>
	<td align="center">
		<input type="text" name="limitsell" value="<%= iting.FTingList(ix).FLimitSell %>" size=6>
	</td>
	<td align="center"><font color="#FF0000"><%= iting.FTingList(ix).FLimitea-iting.FTingList(ix).FLimitSell %></font></td>
	<td align="center">
		<select name="isusing">
			<option value="Y" <% if iting.FTingList(ix).Fisusing="Y" then response.write "selected" %> >전시함</option>
			<option value="N" <% if iting.FTingList(ix).Fisusing="N" then response.write "selected" %> >전시안함</option>
		</select>
	</td>
	<td align="center">
		<select name="sellyn">
			<option value="Y" <% if iting.FTingList(ix).Fsellyn="Y" then response.write "selected" %> >판매함</option>
			<option value="N" <% if iting.FTingList(ix).Fsellyn="N" then response.write "selected" %> >판매안함</option>
		</select>
	</td>
	<td align="center">
		<select name="eventdiv">
			<option value="0" <% if iting.FTingList(ix).Feventdiv="0" then response.write "selected" %> >-</option>
			<option value="1" <% if iting.FTingList(ix).Feventdiv="1" then response.write "selected" %> >이벤트1(상품)</option>
			<option value="2" <% if iting.FTingList(ix).Feventdiv="2" then response.write "selected" %> >이벤트2(기타)</option>
		</select>
	</td>
	<td align="center">
		<input type="text" name="eventcpcode" value="<%= iting.FTingList(ix).FEventCpCode %>" size=7>
	</td>
	<td align="center">
		<input type="button" value="수정" onclick="CheckNEditTing(frm_<%= iting.FTingList(ix).FID %>)">
	</td>
</tr>
</form>
<% next %>
</table>
<%
set iting = Nothing
%>
<br>
<table width="400" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="2">팅 상품추가</td>
</tr>
<form name="frmting" method="post" action="dotingedit.asp">
<input type="hidden" name="mode" value="add">
<tr>
	<td width="100">아이템ID</td>
	<td ><input type="text" name="itemid" value="" size=6></td>
</tr>
<tr>
	<td width="100">구매구분</td>
	<td >
		<select name="userclass">
			<option value="A" >정팅,준팅</option>
			<option value="Y" >정팅</option>
			<option value="N" >정팅,준팅,노팅</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">한정판매</td>
	<td >
		<select name="limitdiv">
			<option value="0" >비한정판매</option>
			<option value="1" >수량한정</option>
			<option value="2" >일별한정</option>
			<option value="3" >월별한정</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">팅포인트(정)</td>
	<td ><input type="text" name="tingpoint" value="" size=7></td>
</tr>
<tr>
	<td width="100">팅포인트(비)</td>
	<td ><input type="text" name="tingpoint_b" value="" size=7></td>
</tr>
<tr>
	<td width="100">한정판매수량</td>
	<td ><input type="text" name="limitea" value="0" size=7></td>
</tr>
<tr>
	<td width="100">현재한정판매수량</td>
	<td ><input type="text" name="limitsell" value="0" size=7></td>
</tr>
<tr>
	<td width="100">전시여부</td>
	<td>
		<select name="isusing">
			<option value="Y">전시함</option>
			<option value="N">전시안함</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">판매여부</td>
	<td>
		<select name="sellyn">
			<option value="Y">판매함</option>
			<option value="N">판매안함</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="button" value="추가" onclick="checkNAddting(frmting)"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->