<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 핑거스 발주서작성 
' History	:  2015.05.27 이상구 생성
'              2017.07.05 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<!-- #include virtual="/academy/lib/classes/order/acabaljucls.asp"-->
<%

Dim FlushCount : FlushCount=100  ''2016/04/18 :: ASP 페이지를 실행하여 Response 버퍼의 구성된 제한이 초과되었습니다.

dim pagesize
dim research
dim yyyy1,mm1,dd1,yyyymmdd,nowdate


'==============================================================================
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)

pagesize = RequestCheckvar(request("pagesize"),10)
research = RequestCheckvar(request("research"),2)


'==============================================================================
if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if (pagesize="") then pagesize=50


dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CAcademyDiyBalju

''총 페이징의 2배 검색
ojumun.FPageSize = pagesize * 3

ojumun.FCurrPage = page

ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1

ojumun.GetBaljuItemListProc


dim ix,iy
dim acabaljucount
acabaljucount = 0

%>
<script language='javascript'>
var acaBaljuCnt = 0;
function CheckNBalju(){
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

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('출고 택배사를 선택 하세요.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('작업 그룹을 선택 하세요.');
		document.all.groupform.workgroup.focus();
		return;
	}
    //C작업장 DAS
    isDasBalju = (document.all.groupform.workgroup.value=="C");

    //DAS 발주 체크, 텐배 150개 이하.
    if ((isDasBalju)&&(tenBaljuCnt>150)){
        alert('DAS 발주는 텐바이텐 배송 150건 미만만 가능합니다. ');
		document.all.groupform.workgroup.focus();
		return;
    }

	var ret = confirm('선택 주문을 새 발주서로 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
		upfrm.workgroup.value = document.all.groupform.workgroup.value;

		if ((upfrm.songjangdiv.value == "91") || (upfrm.songjangdiv.value == "92")) {
			upfrm.songjangdiv.value = "90";
		}

		upfrm.submit();
	}
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnableDiable(icomp){
	//return;
	var frm = document.frm;
	var ischecked = icomp.checked;
	if (ischecked){
		if (icomp.name=="notitemlistinclude"){
			frm.itemlistinclude.checked = !(ischecked);
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
		}

	}

	if (ischecked){
		if (icomp.name=="notbrandlistinclude"){
			frm.brandlistinclude.checked = !(ischecked);
		}else if (icomp.name=="brandlistinclude"){
			frm.notbrandlistinclude.checked = !(ischecked);
		}

	}

	if (icomp.name=="onlyOne"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("/admin/ordermaster/poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function poponebrand(){
	var popwin = window.open("/admin/ordermaster/poponebrand.asp","poponebrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkUpbea(){
    var frm;
    var checkedExists = false;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.tenbeaexists.value!="Y"){
			    frm.cksel.checked = true;
			    AnCheckClick(frm.cksel);
			    checkedExists = true;
			}
		}
	}

	if (checkedExists){
	    document.groupform.songjangdiv.value="24";
	    document.groupform.workgroup.value="Z";
	    CheckNBalju();
	}
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr height="35">
        			<td width="320"><b>기간 : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ 현재</td>
        			<td width="220">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>핑거스배송 건수</b> :
						<select name="pagesize" >
							<option value="10" <% if pagesize="10" then response.write "selected" %> >10</option>
							<option value="20" <% if pagesize="20" then response.write "selected" %> >20</option>
							<option value="50" <% if pagesize="50" then response.write "selected" %> >50</option>
							<option value="200" <% if pagesize="200" then response.write "selected" %> >200</option>
						</select>
        			</td>
					<td>&nbsp;</td>
        		</tr>
        	</table>
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        총 미발주 건수 : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
			총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">총 검색주문건수 : </div>
	        <div id="currtensearchno">핑거스배송 주문건수 : </div>
        </td>
        <td align="right">
		<form name="groupform">
			&nbsp;
		    <select name="songjangdiv">
		        <option value="">택배사선택</option>
               	<option value="4" >CJ택배</option>
		    </select>
			<select name="workgroup">
			   	<option value="">작업그룹</option>
			   	<option value="Y" >Y</option>
		   	</select>
			<input type="button" value="선택주문 발주서작성" onclick="CheckNBalju()" disabled>
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(true)"></td>
	<td width="80">주문번호</td>
	<td width="120">Site</td>
	<td width="50">국가</td>
	<td width="120">UserID</td>
	<td width="120">구매자</td>
	<td width="120">수령인</td>
	<td width="60">결제금액</td>
	<td width="60">구매총액</td>
	<td width="80">결제방법</td>
	<td width="80">거래상태</td>
	<td width="110">주문일</td>
	<td>

    </td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if acabaljucount< CLng(pagesize) then %>
	<% if (ix/FlushCount)=CLNG(ix/FlushCount) then response.write CLNG(ix/FlushCount): response.flush %>
<form name="frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>" method="post" >
<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(ix).FOrderSerial %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sitename" value="<%= ojumun.FItemList(ix).FSiteName %>">
<input type="hidden" name="dlvcontrycode" value="<%= ojumun.FItemList(ix).FDlvcountryCode %>">
<tr align="center" bgcolor="#FFFFFF">
<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(ix).FDlvcountryCode<>"" and ojumun.FItemList(ix).FDlvcountryCode<>"KR","disabled","") %> ></td>
<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
<td><%= ojumun.FItemList(ix).FDlvcountryCode %></td>
<td>
	<%= printUserId(ojumun.FItemList(ix).FUserID, 2, "*") %>
</td>
<td><%= ojumun.FItemList(ix).FBuyName %></td>
<td><%= ojumun.FItemList(ix).FReqName %></td>
<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
<td><%= Left(ojumun.FItemList(ix).FRegDate,16) %></td>
<td>
<% if ojumun.FItemList(ix).Ftenbeaexists then %>
<input type="hidden" name="tenbeaexists" value="Y">
<% acabaljucount = acabaljucount + 1 %>
√ <%= acabaljucount %>
<% else %>
<input type="hidden" name="tenbeaexists" value="N">
<% end if %>
</td>
</tr>
</form>
	<% else %>
		<% exit for %>
	<% end if %>
	<% next %>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<form name="frmArrupdate" method="post" action="dobaljumakerACA.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="sitename" value="">
<input type="hidden" name="songjangdiv" value="">
<input type="hidden" name="workgroup" value="">
<input type="hidden" name="baljutype" value="">
<input type="hidden" name="extSiteName" value="diyitem">
</form>
<%
set ojumun = Nothing
%>

<script language='javascript'>
document.all.currsearchno.innerHTML = "검색갯수 : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "핑거스배송 검색갯수 : <Font color='#3333FF'><%= acabaljucount %></font>";
acaBaljuCnt = 1*<%= acabaljucount %>;
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
