<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<%

dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, itemid
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim sitename

searchfield = RequestCheckvar(request("searchfield"),16)
userid = RequestCheckvar(request("userid"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
username = RequestCheckvar(request("username"),16)
userhp = RequestCheckvar(request("userhp"),16)
etcfield = RequestCheckvar(request("etcfield"),2)
etcstring = RequestCheckvar(request("etcstring"),32)
itemid = RequestCheckvar(request("itemid"),10)

checkYYYYMMDD = RequestCheckvar(request("checkYYYYMMDD"),1)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

sitename = RequestCheckvar(request("sitename"),16)


'==============================================================================
dim nowdate, searchnextdate

if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

sitename = "academy"
'==============================================================================

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CRequestLecture
ojumun.FPageSize = 10
ojumun.FCurrPage = page

ojumun.FRectSiteName = sitename




if checkYYYYMMDD="Y" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if


if (searchfield = "orderserial") then
        '주문번호
        ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
        '고객아이디
        ojumun.FRectUserID = userid
        ojumun.FPageSize = 50
elseif (searchfield = "itemid") then
        '강좌아이디
        ojumun.FRectItemID = itemid
        ojumun.FPageSize = 50
elseif (searchfield = "username") then
        '구매자명
        ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
        '구매자핸드폰
        ojumun.FRectBuyHp = userhp
        ojumun.FPageSize = 50
elseif (searchfield = "etcfield") then
        '기타조건
        if etcfield="01" then
        	ojumun.FRectBuyname = etcstring
        elseif etcfield="02" then
        	ojumun.FRectReqName = etcstring
        elseif etcfield="03" then
        	ojumun.FRectUserID = etcstring
        elseif etcfield="04" then
        	ojumun.FRectIpkumName = etcstring
        elseif etcfield="06" then
        	ojumun.FRectSubTotalPrice = etcstring
        elseif etcfield="07" then
        	ojumun.FRectBuyHp = etcstring
        elseif etcfield="08" then
        	ojumun.FRectReqHp = etcstring
        end if
end if

ojumun.GetRequestLectureMasterList

dim ix,i
dim totalavailcount

%>

<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','LecOrderDetail');
    frm.target = 'lec_orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function GotoOrderDetail(orderserial) {
        top.mainframe.location.href = "lec_request_detail.asp?orderserial=" + orderserial;
        // var popwin = window.open('lec_orderdetail.asp?orderserial=' + orderserial,'LecOrderDetail','width=800,height=400,scrollbars=yes,resizable=yes');
        // popwin.focus();
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

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function ChangeCheckbox(frmname, frmvalue) {
        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        if ((frm.elements[i].name == frmname) && (frm.elements[i].value == frmvalue)) {
                                frm.elements[i].checked = true;
                        }
                }
        }
}

function FocusAndSelect(frm, obj){
        ChangeFormBgColor(frm);

        obj.focus();
        obj.select();
}

function ChangeFormBgColor(frm) {
        // style='background-color:#DDDDFF'
        var radioselected = false;
        var checkboxchecked = false;
        var ischecked = false;

        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "checkbox") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "text") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFFF";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }

                if (frm.elements[i].type == "select-one") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFFF";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }
        }
}

// tr 색상변경
var pre_selected_row = null;
var pre_selected_row_color = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color != null) {
	        pre_selected_row.bgColor = pre_selected_row_color;
        }
        pre_selected_row = e;
        pre_selected_row_color = defcolor;
        e.bgColor = selcolor;
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	<input type="radio" name="searchfield" value="orderserial" <% if searchfield="orderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.orderserial)"> 주문번호
    		<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'orderserial'); FocusAndSelect(frm, frm.orderserial);">

    		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> 아이디
    		<input type="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

    		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> 구매자명
    		<input type="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

    		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> 구매자핸드폰
    		<input type="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">
			<br>
			<!--
			<input type="radio" name="searchfield" value="itemid" <% if searchfield="itemid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.itemid)"> 강좌번호
    		<input type="text" name="itemid" value="<%= itemid %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'itemid'); FocusAndSelect(frm, frm.itemid);">
			-->

            <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> 기타조건
    		<select name="etcfield">
    		  <option value="">선택</option>
                  <option value="01" <% if etcfield="01" then response.write "selected" %> >구매자 명</option>
                  <!--
                  <option value="02" <% if etcfield="02" then response.write "selected" %> >수령인 명</option>
                  -->
                  <!--
                  <option value="03" <% if etcfield="03" then response.write "selected" %> >아이디</option>
                  -->
                  <option value="04" <% if etcfield="04" then response.write "selected" %> >입금자 명</option>
                  <option value="06" <% if etcfield="06" then response.write "selected" %> >결제금액</option>
                  <option value="07" <% if etcfield="07" then response.write "selected" %> >구매자 핸드폰</option>
                  <!--
                  <option value="08" <% if etcfield="08" then response.write "selected" %> >수령인 핸드폰</option>
                  -->
                  <!--
                  <option value="09" <% if etcfield="09" then response.write "selected" %> >송장번호</option>
                	-->
                </select>
    		<input type="text" name="etcstring" value="<%= etcstring %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
    		<br>
    		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td width="40">구분</td>
	    	<td width="80">주문번호</td>
	    	<td width="50">거래상태</td>
	    	<td width="60">결제방법</td>
	    	<td width="70">총결제금액</td>
	    	<td>UserID</td>
	    	<td width="40">수강생</td>
	    	<td width="60">신청자</td>
	    	<td width="90">신청자HP</td>
	    	<td width="100">강좌명 / 상품명</td>
	    	<td width="70">신청일</td>
	    	<td width="70">입금일</td>
	    </tr>
<% if ojumun.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="14" align="center">[검색결과가 없습니다.]</td>
	    </tr>
<% end if %>
<% if ojumun.FresultCount > 0 then %>
    <% for ix=0 to ojumun.FresultCount-1 %>
		<% if ojumun.FItemList(ix).IsAvailable then %>
		    <% totalavailcount = totalavailcount + ojumun.FItemList(ix).Ftotalitemno %>
		<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); " style="cursor:hand">
		<% else %>
		<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); " style="cursor:hand">
		<% end if %>
			<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
			<td><a href="javascript:GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
			<td><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></acronym></td>
			<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
			<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
			<td align="left"><a href="?searchfield=userid&userid=<%= ojumun.FItemList(ix).FUserID %>"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></a></font></td>
			<td><%= ojumun.FItemList(ix).Ftotalitemno %></td>
			<td><%= ojumun.FItemList(ix).FBuyName %></td>
			<td><%= ojumun.FItemList(ix).Fbuyhp %></td>
			<td align="left"><acronym title="<%= ojumun.FItemList(ix).Fgoodsnames %>"><%= DdotFormat(ojumun.FItemList(ix).Fgoodsnames,18) %></acronym></td>
			<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td colspan="6"></td>
			<td align="center"><%= totalavailcount %></td>
			<td colspan="6"></td>
		</tr>
	</table>
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
    			<% if ix>ojumun.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ojumun.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<%
set ojumun = Nothing
%>
<script language='javascript'>
ChangeFormBgColor(frm);
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->