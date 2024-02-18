<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2015.11.05 한용민 생성
'	Description : 포장 서비스
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/pack_cls.asp"-->

<%
Dim orderserial, i, midx, userid, tmpmidx
	orderserial = requestcheckvar(request("orderserial"),11)
	midx = getNumeric(requestcheckvar(request("midx"),10))

dim vtitle, vmessage

if orderserial="" then
	response.write "<script type='text/javascript'>alert('주문번호가 없습니다.'); self.close();</script>"
	dbget.close()	:	response.end
end if

dim cpacksum
set cpacksum = new Cpack
	cpacksum.FRectOrderSerial = orderserial
	cpacksum.Getpojang_itemlist()

%>

<script type="text/javascript">

function detailview(orderserial, midx){
	location.replace("/cscenter/pack/pojang_view.asp?orderserial="+orderserial+"&midx="+midx);
}

function editproc(midx){
	if (midx==''){
		alert('일렬번호가 없습니다.');
		return;
	}

	if (pojangfrm.title.value == '' || GetByteLength(pojangfrm.title.value) > 60){
		alert("선물 포장명이 없거나 제한길이를 초과하였습니다. 60자 까지 작성 가능합니다.");
		pojangfrm.title.focus();
		return;
	}
	if (pojangfrm.message.value != '' && GetByteLength(pojangfrm.title.value) > 100){
		alert("선물 메세지가 제한길이를 초과하였습니다. 100자 까지 작성 가능합니다.");
		pojangfrm.message.focus();
		return;
	}

	pojangfrm.mode.value='editpojang';
	pojangfrm.midx.value=midx;
	pojangfrm.action = "/cscenter/pack/pojang_process.asp";
	pojangfrm.submit();
	return;
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		주문번호 : <%= orderserial %>
	</td>
	<td align="right">
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</tr>
</table>

<br>
<font color="red">※선물포장내역</font>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		총박스수 : <b><%= cpacksum.Fpackcnt %></b> / 총상품수량합계 : <b><%= cpacksum.Fpackitemcnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>포장명</td>
	<td>포장메세지</td>
	<td>상품수<br>합계</td>
	<td>포장일</td>
	<td>삭제여부</td>
	<td>박스번호<br>(고객오픈 X)</td>
	<td>비고</td>
</tr>
<% if cpacksum.FresultCount>0 then %>
	<% for i=0 to cpacksum.FresultCount-1 %>
		<% if cstr(midx)=cstr(cpacksum.FItemList(i).fmidx) then %>
			<%
			vtitle = cpacksum.FItemList(i).ftitle
			vmessage = cpacksum.FItemList(i).fmessage
			%>
			<tr align="center" bgcolor="orange" >
		<% else %>
			<tr align="center" bgcolor="#FFFFFF" >
		<% end if %>

		<% if tmpmidx="" or cstr(tmpmidx)<>cstr(cpacksum.FItemList(i).fmidx)  then %>
			<td>
				<%= chrbyte(cpacksum.FItemList(i).ftitle,10,"Y") %>
			</td>
			<td>
				<%= chrbyte(cpacksum.FItemList(i).fmessage,10,"Y") %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fpackitemcnt %>
			</td>
			<td>
				<%= FormatDate(cpacksum.FItemList(i).fregdate,"0000.00.00") %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fcancelyn %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fmidx %>
			</td>
			<td>
				<input type="button" onclick="detailview('<%= orderserial %>','<%= cpacksum.FItemList(i).fmidx %>');" value="수정" class="button">
			</td>
		<% end if %>

		<% tmpmidx = cpacksum.FItemList(i).fmidx %>

		<% if cstr(tmpmidx)=cstr(cpacksum.FItemList(i).fmidx) then %>
			</tr>
			<tr>
				<td colspan=7 align="right" bgcolor="#FFFFFF">
					<table width="900" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr bgcolor="#FFFFFF" >
						<td width=55>
							<img src="<%= cpacksum.FItemList(i).FImageSmall %>" width=50 height=50>
						</td>
						<td width=65>
							상품코드:
							<br><%= cpacksum.FItemList(i).FItemID %>
						</td>
						<td width=130>
							브랜드명:
							<br><%= cpacksum.FItemList(i).FBrandName %>
						</td>
						<td>
							상품명:
							<br><%= cpacksum.FItemList(i).FItemName %>
						</td>
						<td width=170>
							<% if cpacksum.FItemList(i).FItemOptionName<>"" then %>
								옵션명:
								<br><%= cpacksum.FItemList(i).FItemOptionName %>
							<% end if %>
						</td>
						<td width=70>
							수량: <%= cpacksum.FItemList(i).FItemEa %>
						</td>
						<td width=50>
							삭제: <%= cpacksum.FItemList(i).fcancelyn %>
						</td>
					</tr>
					</table>
				</td>
		<% end if %>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% if midx<>"" then %>
	<br>
	<font color="red">※선물포장수정</font>
	<br>
	<form name="pojangfrm" method="post" action="" style="margin:0px;">
	<input type="hidden" name="mode">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<tr>
		<td bgcolor="#e1e1e1" align="center">박스번호</td>
		<td bgcolor="#FFFFFF">
			<%= midx %>
			<input type="hidden" name="midx">
		</td>
	</tr>
	<tr>
		<td bgcolor="#e1e1e1" align="center">포장명</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="title" value="<%= vtitle %>" size="100">
		</td>
	</tr>
	<tr>
		<td bgcolor="#e1e1e1" align="center">포장메세지</td>
		<td bgcolor="#FFFFFF">
			<textarea name="message" style="width:600px;" rows="5"><%= vmessage %></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" align="center" colspan=2>
			<input type="button" onclick="editproc('<%= midx %>');" value="수정" class="button">
		</td>
	</tr>
	</table>
	</form>
<% end if %>

<%
set cpacksum = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
