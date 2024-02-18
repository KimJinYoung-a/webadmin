<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim idx,mode
dim olec

idx = request("idx")
mode = request("mode")

if idx="" then idx=0
set olec = new CLectureDetail
olec.GetLectureDetail idx

dim itemid,odetail
itemid = olec.Flinkitemid
set odetail = new CLecture
odetail.FRectItemID = itemid
odetail.GetLectureRegList

dim i
dim totno

totno =0
%>
<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function SendMsg(){
	if (!CheckSelected()){
		alert('하나이상 선택 하셔야 합니다.');
		return;
	}

	var ret = confirm('메시지를 보내시겠습니까?');
	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					refreshFrm.orderserial.value = refreshFrm.orderserial.value + frm.orderserial.value + ",";
				}
			}
		}
		var popwin = window.open('','refreshFrm','width=300 height=300');
		popwin.focus();
		refreshFrm.idx.value="<%=idx %>";
		refreshFrm.target = "refreshFrm";
		refreshFrm.action = "/admin/lecture/lecture_inputmsg.asp";
		refreshFrm.submit();
	}
}
</script>


<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >강좌명</td>
	<td bgcolor="#FFFFFF"><% = olec.Flectitle %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><% = olec.Flecturer %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌비</td>
	<td bgcolor="#FFFFFF">
		<% = olec.Flecsum %>
		<% if olec.Fmatinclude = "Y" then %>
		(재료비포함)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료비</td>
	<td bgcolor="#FFFFFF"><% = olec.Fmatsum %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의기간<br>(주기)</td>
	<td bgcolor="#FFFFFF"><% = olec.Flecperiod %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의시간</td>
	<td bgcolor="#FFFFFF"><% = olec.Flectime %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>강좌일시</td>
	<td bgcolor="#FFFFFF">
		<% if Left(olec.Flecdate01,10)<>"1900-01-01" then %>
		1주 : <% = olec.Flecdate01 %>~<% = olec.Flecdate01_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate02,10)<>"1900-01-01" then %>
		2주 : <% = olec.Flecdate02 %>~<% = olec.Flecdate02_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate03,10)<>"1900-01-01" then %>
		3주 : <% = olec.Flecdate03 %>~<% = olec.Flecdate03_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate04,10)<>"1900-01-01" then %>
		4주 : <% = olec.Flecdate04 %>~<% = olec.Flecdate04_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate05,10)<>"1900-01-01" then %>
		5주 : <% = olec.Flecdate05 %>~<% = olec.Flecdate05_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate06,10)<>"1900-01-01" then %>
		6주 : <% = olec.Flecdate06 %>~<% = olec.Flecdate06_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate07,10)<>"1900-01-01" then %>
		7주 : <% = olec.Flecdate07 %>~<% = olec.Flecdate07_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate08,10)<>"1900-01-01" then %>
		8주 : <% = olec.Flecdate08 %>~<% = olec.Flecdate08_end %><br>
		<% end if %>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<tr>
		<td align=right class="a" bgcolor="#FFFFFF">
		<input type="button" value="메시지 보내기" onClick="SendMsg();" class="button">
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>주문번호</td>
	<td>상태</td>
	<td>성명</td>
	<td>아이디</td>
	<td>수량</td>
	<td>전화</td>
	<td>핸드폰</td>
	<td>이메일</td>
	<td>주문일</td>
	<td>결제일</td>
</tr>
<% for i=0 to odetail.FResultCount -1 %>
<%
if Not odetail.FItemList(i).IsCancel then
totno = totno + odetail.FItemList(i).Fitemno
end if
%>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name=orderserial value=<%= odetail.FItemList(i).FOrderserial %>>
<tr bgcolor="#FFFFFF">
	<td>
		<% if odetail.FItemList(i).FIpkumdiv >=3 and Not(odetail.FItemList(i).IsCancel) then %>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<% else %>
		<input type="checkbox" name="cksel" disabled>
		<% end if %>
	</td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FOrderserial %></font></td>
	<td><Font color=<%= odetail.FItemList(i).GetStateColor %> ><%= odetail.FItemList(i).GetStateName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserID %></font></td>
	<td align="center"><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).Fitemno %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyPhone %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyHp %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserEmail %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FRegdate %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FIpkumDate %></font></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan=5></td>
	<td align="center"><%= totno %></td>
	<td colspan=6></td>
</tr>
</table>
<form name=refreshFrm method=post>
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->