<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 b2c 매출 작성
' History : 2012.08.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim shopid, startdate, enddate, onlymifinish, research , menupos ,b2ccharge
dim nowdate, yyyy1, yyyy2, mm1, mm2, dd1, dd2, i, ttlsell, ttlsuply, ttlbuy
	shopid = requestCheckVar(request("shopid"),32)
	onlymifinish = requestCheckVar(request("onlymifinish"),2)
	research = requestCheckVar(request("research"),2)
	menupos = requestCheckVar(request("menupos"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)

if (research="") and (onlymifinish="") then onlymifinish="on"

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)

	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2) - 3
	dd1   = "01" ''Mid(nowdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

startdate = CStr(DateSerial(yyyy1 , mm1 , dd1))
enddate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim ob2c
set ob2c = new CFranjungsan
	ob2c.FRectshopid = shopid
	ob2c.FRectStartDate = startdate
	ob2c.FRectEndDate = enddate
	ob2c.FRectonlymifinish = onlymifinish
	
	if shopid<>"" then
		ob2c.getB2Cmaechullist
		
		b2ccharge = "15"
	else	
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장을 선택하신 후 검색하세요.');"		
		response.write "</script>"		
	end if

ttlsell = 0
ttlsuply = 0
ttlbuy = 0

if b2ccharge = "" then b2ccharge = "15"
%>

<script type='text/javascript'>

function totalCheck(){
	var f = document.frmArr;
	var objStr = "check";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}
	
function frmsubmit(){
	frm.submit();
}

function editOffDesinger(shopid,designerid){
	var editOffDesinger = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"editOffDesinger","width=640,height=540,scrollbars=yes,resizable=yes");
	editOffDesinger.focus();
}

function popjumundetail(yyyy1,mm1,dd1,shopid){
	var popjumundetail = window.open('/admin/offshop/todaysellmaster.asp?menupos=<%= menupos %>&datefg=jumun&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid,'popjumundetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popjumundetail.focus();
}

function popitemdetail(yyyy1,mm1,dd1,shopid){
	var popitemdetail = window.open('/admin/offshop/todayselldetail.asp?menupos=<%= menupos %>&datefg=jumun&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid,'popitemdetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popitemdetail.focus();
}

function SaveArr(){

	var frmmaster = document.frm;
	var frm = document.frmArr;
		
	if (frmmaster.b2ccharge.value==''){
		alert('B2C 수수료가 없습니다');
		frmmaster.b2ccharge.focus();
	}
	
	frm.b2ccharge.value = frmmaster.b2ccharge.value;
	
	var ret = 0;
    for (i=0; i< document.getElementsByName("check").length; i++)
    {
        if (document.getElementsByName("check")[i].checked == true)
        {
            ret = ret + 1;
        }
    }
	if (ret == 0)
	{
		alert("선택값이 없습니다.");
		return;
	}

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>		
				매장 :
				<% drawSelectBoxOffShopNot000 "shopid",shopid %>
				날짜 :
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<Br><input type=checkbox name=onlymifinish <% if onlymifinish="on" then response.write "checked" %> >미처리 내역만		
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>	
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    	<% if shopid <> "" then %>
    		B2C수수료 : 매출의 <input type="text" name="b2ccharge" value="<%= b2ccharge %>" size=4 maxlength=5>%
    	<% end if %>
    </td>
    <td align="right">
    </td>
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmArr" method="post" action="/admin/offshop/domeaipchulgojungsan.asp">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="workidx" value="">
<input type="hidden" name="mode" value="b2cmaechul">
<input type="hidden" name="b2ccharge">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=ob2c.FresultCount%></b>&nbsp;&nbsp;<% if ob2c.FresultCount = "400" then response.write "최대 1000건" %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="totalCheck()"></td>
	<td>매출일</td>
	<td>매장</td>
	<td>총건수</td>
	<td>총매출액</td>
	<td>
		총업체매입가
	</td>
	<td>총매장공급가</td>
	<td>기처리</td>
	<td>비고</td>
</tr>
<%
if ob2c.FresultCount > 0 then
	
for i=0 to ob2c.FResultCount-1

ttlsell = ttlsell + ob2c.FItemList(i).Ftotsum
ttlsuply = ttlsuply + ob2c.FItemList(i).Frealjungsansum
ttlbuy = ttlbuy + ob2c.FItemList(i).fbuyprice

%>
<tr align="center" bgcolor="ffffff">
	<td>
		<input type="checkbox" name="check" value="'<%= ob2c.FItemList(i).fyyyymmdd %>'" onClick="AnCheckClick(this);">
	</td>
	<td>
		<%= ob2c.FItemList(i).fyyyymmdd %>
	</td>
	<td>
		<%= ob2c.FItemList(i).Fshopid %>
	</td>
	<td><%= ob2c.FItemList(i).Ftotitemcnt %></td>
	<td align="right">
		<%= FormatNumber(ob2c.FItemList(i).Ftotsum,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(ob2c.FItemList(i).Frealjungsansum,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(ob2c.FItemList(i).fbuyprice,0) %>
	</td>
	<td><%= ob2c.FItemList(i).Fprecheckidx %></td>
	<td>
		<input type="button" onclick="popitemdetail('<%= left(ob2c.FItemList(i).fyyyymmdd,4) %>','<%= mid(ob2c.FItemList(i).fyyyymmdd,6,2) %>','<%= right(ob2c.FItemList(i).fyyyymmdd,2) %>','<%= ob2c.FItemList(i).FShopid %>');" value="상품상세" class="button">
		
		<% if not(C_IS_Maker_Upche) then %> 
			<input type="button" onclick="popjumundetail('<%= left(ob2c.FItemList(i).fyyyymmdd,4) %>','<%= mid(ob2c.FItemList(i).fyyyymmdd,6,2) %>','<%= right(ob2c.FItemList(i).fyyyymmdd,2) %>','<%= ob2c.FItemList(i).FShopid %>');" value="주문상세" class="button">
		<% end if %>
	</td>
</tr>
<% Next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=4>합계</td>
	<td align="right"><%= formatnumber(ttlsell,0) %></td>
	<td align="right"><%= formatnumber(ttlsuply,0) %></td>
	<td align="right"><%= formatnumber(ttlbuy,0) %></td>
	<td ></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center"><input type="button" value="선택내역저장" onclick="SaveArr()" class="button_s"></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</form>
</table>

<%
set ob2c = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
