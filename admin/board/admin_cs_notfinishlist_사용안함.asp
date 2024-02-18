<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/afterservicecls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim notfinish,research,onlysongjang
research = request("research")
notfinish = request("notfinish")
onlysongjang = request("onlysongjang")


if (research="") and (notfinish="") then
	notfinish="on"
end if

if (research="") and (onlysongjang="") then
	onlysongjang="on"
end if

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	date1 = dateAdd("d",-10,nowdate)
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(CDate(yyyy2 + "-" + mm2 + "-" + dd2)+1),10)
end if

dim ioneas,i
set ioneas = new CAfterService

ioneas.FRectOnlyJupsu = notfinish
ioneas.FRectOnlyBeasong = "on"
ioneas.FRectStartDay = yyyy1 + "-" + mm1 + "-" + dd1
ioneas.FRectEndDay = Edate
ioneas.GetAsList
%>
<script language='javascript'>
function ShowSongJangDetail(frm){
	frm.submit();
}

function AnCheckNSongjangView(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 내역으로 송장파일을 작성하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idarr.value = upfrm.idarr.value + "|" + frm.id.value;
				}
			}
		}
		upfrm.submit();
	}

	upfrm.target = 'popsongjangmaker';
    upfrm.action="/admin/etcsongjang/popsongjangmaker.asp"
	upfrm.submit();
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="T">
	<tr>
		<td class="a" width="460">
		조회기간 :
			<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="notfinish" <% if notfinish="on" then response.write "checked" %> >접수상태만 검색
		</td>
		<td class="a" align="right">
			<a href="javascript:ShowSongJangDetail(frm)"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr>
	<td colspan="9"  align="right" bgcolor="#FFFFFF"><a href="javascript:window.print();">print</a></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="50">구분</td>
	<td width="30">상태</td>
	<td width="70">주문번호</td>
	<td width="40">고객명</td>
	<td width="60">고객ID</td>
	<td width="40">등록자</td>
	<td width="60">등록일</td>
	<td width="120">제목</td>
	<td width="200">내용</td>
</tr>
<% for i=0 to ioneas.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= CsGubun2Name(ioneas.FAsItemList(i).FDivCD) %></td>
	<td><%= CsState2Name(ioneas.FAsItemList(i).FCurrstate) %></td>
	<td><%= ioneas.FAsItemList(i).FOrderSerial %></td>
	<td><%= ioneas.FAsItemList(i).FCustomerName %></td>
	<td><%= ioneas.FAsItemList(i).FUserID %></td>
	<td><%= ioneas.FAsItemList(i).Fwritename %></td>
	<td><%= Left(CStr(ioneas.FAsItemList(i).Fregdate),10) %></td>
	<td><%= ioneas.FAsItemList(i).FTitle %></td>
	<td><%= ioneas.FAsItemList(i).Fcontents_jupsu %></td>
</tr>
<% next %>
</table>
<%
set ioneas = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->