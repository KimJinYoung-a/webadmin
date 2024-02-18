<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  목표매출
' History : 2013.03.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/targetmaechul/targetmaechul_cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->
<%
Dim page ,ctarget , yyyy ,i , shopid ,research ,gubuntype ,gubun
	page = requestcheckvar(request("page"),10)
	gubun = requestcheckvar(request("gubun"),10)
	gubuntype = requestcheckvar(request("gubuntype"),10)
	research = requestcheckvar(request("research"),2)
	shopid = requestcheckvar(request("shopid"),32)
	yyyy = requestcheckvar(request("yyyy"),4)
	
	if yyyy = "" then yyyy = year(now())
	if page = "" then page = 1
	if gubuntype = "" then gubuntype = "1"
	if gubuntype = "1" then gubun = "0"	'//조닝별목표매출만 구분값 있음
		
if research <> "ON" and shopid = "" then
	'/매장
	if (C_IS_SHOP) then
		
		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	end if
end if
	
set ctarget = new ctargetmaechul_list
	ctarget.FRectyyyy = yyyy
	ctarget.frectshopid = shopid
	ctarget.frectgubuntype = gubuntype
	ctarget.frectgubun = gubun
	
	if shopid <> "" then
		
		'/조닝별목표매출일 경우
		if gubuntype = "2" then
			if gubun <> "" and gubun <> "0" then
				ctarget.gettargetmaechul
			else
				response.write "<script language='javascript'>"
				response.write "	alert('조닝을 선택 하세요');"
				response.write "</script>"
			end if
		else
			ctarget.gettargetmaechul
		end if
	else
		response.write "<script language='javascript'>"
		response.write "	alert('매장을 선택해 주세요');"
		response.write "</script>"
	end if
%>

<script language="javascript">

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

//선택상품 작년대비 목표매출 계산
function chmaechul(){

	var searchfrm = document.frm;
	
	if (!IsDigit(searchfrm.maechulpro.value)){
		alert('목표대비 %로 숫자만 입력 가능합니다.');
		searchfrm.maechulpro.focus();
		return;
	}
	
	if (searchfrm.maechulpro.value<1){
		alert('목표대비 %로 정확히 입력하세요.');
		searchfrm.maechulpro.focus();
		return;
	}
	
	var frm;
	var pass = false;
	
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}
				
	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.targetmaechul.value = Math.round(frm.realsellsum.value * (searchfrm.maechulpro.value/100));
			}
		}
	}
}

//선택상품 저장
function saveArr(){

	var searchfrm = document.frm;
	
	if (searchfrm.shopid.value==''){
		alert('매장이 선택되지 않았습니다');
		return;
	}
	
	var frm;
	var pass = false;
	
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}
				
	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	frmarr.mode.value = "";
	frmarr.solar_date.value = "";
	frmarr.yyyymm.value = "";
	frmarr.shopid.value = "";
	frmarr.gubuntype.value = "";
	frmarr.gubun.value ="";
	frmarr.targetmaechul.value ="";
	 
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsDigit(frm.targetmaechul.value)){
					alert('목표매출은 숫자만 가능합니다.');
					frm.targetmaechul.focus();
					return;
				}
				
				if (frm.targetmaechul.value<1){
					alert('목표매출을 정확히 입력하세요.');
					frm.targetmaechul.focus();
					return;
				}
				
				frmarr.yyyymm.value = frmarr.yyyymm.value + frm.yyyymm.value + ","				
				frmarr.solar_date.value = frmarr.solar_date.value + frm.solar_date.value + ","
				frmarr.targetmaechul.value = frmarr.targetmaechul.value + frm.targetmaechul.value + ","

			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frmarr.mode.value = 'targetreg';
		frmarr.shopid.value = '<%=shopid%>';
		frmarr.gubuntype.value = searchfrm.gubuntype.value;
		frmarr.gubun.value = searchfrm.gubun.value;
		frmarr.submit();
	}
}

function frmsubmit(){
	frm.submit();
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
     	목표구분 :
		<select name="gubuntype">
			<% fnOptCommonCode "gubuntype",gubuntype %>
		</select>
		<% if gubuntype = "2" then %>
			<% Call zoneselectbox(shopid,"gubun",gubun,"") %>
		<% else %>
			<input type="hidden" name="gubun" value="0">
		<% end if %>
		
		<br>매장 : <% drawSelectBoxOffShopdiv_off "shopid" , shopid ,"1,3,7","Y","" %>
		년도 : <% DrawyearBoxdynamic "yyyy",yyyy,"" %>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
</table>
<!---- /검색 ---->

<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr>
	<td align="left">
		<% IF ctarget.fresultcount > 0 THEN %>
			목표매출 = 작년매출대비 <input type="text" name="maechulpro" value="0" size=5 maxlength=6>%
			<input type="button" value="선택계산" onClick="chmaechul();" class="button">
		<% end if %>		
	</td>
	<td align="right">
		<% IF ctarget.fresultcount > 0 THEN %>
			<input type="button" value="선택수정" onClick="saveArr()" class="button">
		<% end if %>
	</td>	
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="left">
		검색결과 : <b><%=ctarget.ftotalcount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=40>
		<input type="checkbox" name="ck_all" onclick="SelectCk(this)">
	</td>    				    				    	
	<td>날짜</td>
	<td><%=yyyy-1%>년<br>매출</td>
	<td>목표<br>매출</td>
	<td>최종<br>수정</td>
</tr>	
<% 
IF ctarget.fresultcount > 0 THEN
	
For i = 0 To ctarget.fresultcount -1
%>
<form name="frmBuyPrc_<%=i%>" method="get">			
<input type="hidden" name="solar_date" value="<%= ctarget.FItemList(i).fsolar_date %>">
<input type="hidden" name="yyyymm" value="<%= ctarget.FItemList(i).fyyyymm %>">
<input type="hidden" name="gubuntype" value="<%= ctarget.FItemList(i).fgubuntype %>">
<input type="hidden" name="gubun" value="<%= ctarget.FItemList(i).fgubun %>">
<input type="hidden" name="realsellsum" value="<%= ctarget.FItemList(i).frealsellsum %>">
<tr align="center" bgcolor=<%IF ctarget.FItemList(i).fyyyymm = "" or isnull(ctarget.FItemList(i).fyyyymm) THEN%>"#f1f1f1"<%ELSE%>"#FFFFFF"<%END IF%>> 			
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>    				    	
    <td>
    	<%= ctarget.FItemList(i).fsolar_date %>
    </td>
    <td>
    	<%= FormatNumber(ctarget.FItemList(i).frealsellsum,0) %>
    </td>      
    <td>
    	<input type="text" name="targetmaechul" onKeyup="CheckThis(frmBuyPrc_<%= i %>);" value="<%= ctarget.FItemList(i).ftargetmaechul %>" size="12" maxlength="13" style="text-align:right;">
    </td>
    <td>
    	<%= ctarget.FItemList(i).flastadminid %>
    	<Br><%= ctarget.FItemList(i).flastupdate %>
    </td>	    
</tr>
</form>
<% next %>

<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
</tr>	
<% END IF %>
<form name="frmarr" method="post" action="/common/offshop/maechul/targetmaechul/targetmaechul_process.asp">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="solar_date">
	<input type="hidden" name="yyyymm">
	<input type="hidden" name="shopid">
	<input type="hidden" name="gubuntype">
	<input type="hidden" name="gubun">
	<input type="hidden" name="targetmaechul">
</form>		    
</table>

<%
set ctarget = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->