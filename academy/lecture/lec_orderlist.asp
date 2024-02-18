<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 강좌 관리
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, itemid
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite ,yyyy1,yyyy2,mm1,mm2,dd1,dd2 ,jumundiv, jumunsite ,lecOption
dim ix,i ,totalavailcount ,olecture ,oLectOption
	searchfield = RequestCheckvar(request("searchfield"),16)
	userid      = RequestCheckvar(request("userid"),32)
	orderserial = RequestCheckvar(request("orderserial"),16)
	username    = RequestCheckvar(request("username"),16)
	userhp      = RequestCheckvar(request("userhp"),16)
	etcfield    = RequestCheckvar(request("etcfield"),2)
	etcstring   = RequestCheckvar(request("etcstring"),32)
	itemid      = RequestCheckvar(request("itemid"),10)
	lecOption   = RequestCheckvar(request("lecOption"),4)
	checkYYYYMMDD = RequestCheckvar(request("checkYYYYMMDD"),1)
	checkJumunDiv = RequestCheckvar(request("checkJumunDiv"),1)
	checkJumunSite = RequestCheckvar(request("checkJumunSite"),1)
	yyyy1 = RequestCheckvar(request("yyyy1"),4)
	mm1 = RequestCheckvar(request("mm1"),2)
	dd1 = RequestCheckvar(request("dd1"),2)
	yyyy2 = RequestCheckvar(request("yyyy2"),4)
	mm2 = RequestCheckvar(request("mm2"),2)
	dd2 = RequestCheckvar(request("dd2"),2)
	jumundiv = RequestCheckvar(request("jumundiv"),10)
	page = RequestCheckvar(request("page"),10)
	if (page="") then page=1
	
'==============================================================================
dim nowdate, searchnextdate ,page ,ojumun

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
'==============================================================================

set ojumun = new CLectureFingerOrder
	ojumun.FPageSize = 200
	ojumun.FCurrPage = page
	
	if checkYYYYMMDD="Y" then
		ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
		ojumun.FRectRegEnd = searchnextdate
	end if

	if (checkJumunDiv = "Y") then
	        if (jumundiv="flowers") then
	        	ojumun.FRectIsFlower = "Y"
	        elseif (jumundiv="lecture") then
	                ojumun.FRectIsLecture = "Y"
	        elseif (jumundiv="minus") then
	                ojumun.FRectIsMinus = "Y"
	        end if
	end if
	
	if (checkJumunSite = "Y") then
		ojumun.FRectExtSiteName = jumunsite
	end if
	
	if (searchfield = "orderserial") then
	    '주문번호
	    ojumun.FRectOrderSerial = orderserial
	elseif (searchfield = "userid") then
	    '고객아이디
	    ojumun.FRectUserID = userid
	elseif (searchfield = "username") then
	    '구매자명
	    ojumun.FRectBuyname = username
	elseif (searchfield = "userhp") then
	    '구매자핸드폰
	    ojumun.FRectBuyHp = userhp
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
	    elseif etcfield="09" then
	    	ojumun.FRectReqSongjangNo = etcstring
	    end if
	end if
	
	if (searchfield = "itemid") then
		ojumun.FRectItemID = itemid
		ojumun.FREctItemOption=lecOption
		ojumun.GetFingerOrderListByItemID
	else
		ojumun.GetFingerOrderList
	end if
	
	set olecture = new CLecture
		olecture.FRectIdx = itemid
	
	if (searchfield = "itemid") then
		olecture.GetOneLecture
	end if

'// 옵션정보
Set oLectOption = New CLectOption
	oLectOption.FRectidx = itemid
	''oLectOption.FRectOptIsUsing = "Y"
	if itemid<>"" then
		oLectOption.GetLectOptionInfo
	end if

dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = itemid
'olecschedule.FRectOptCd = lecOption

if (searchfield = "itemid") then
	olecschedule.GetOneLecSchedule
end if
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
    var popwin = window.open('/cscenterv2/lecture/lecturedetail_view.asp?orderserial=' + orderserial,'LecOrderDetail','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ViewUserInfo(frm){
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

function NewWindow(v){
  var p = (v);
  w = window.open("http://www.thefingers.co.kr/photo_album/pop_photo.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
  w.focus();
}

//자바스크립트 동적 제어를 위해 아이프레임으로 넘긴다
function SongJangPrintProc_Fingers(OrderSerial,startdate,entryname,lec_title,SubTotalPrice,barcodelecprice,barcodematprice,itemid){
	viewfrm.OrderSerial.value='';
	viewfrm.startdate.value='';
	viewfrm.entryname.value='';
	viewfrm.lec_title.value='';
	viewfrm.SubTotalPrice.value='';
	viewfrm.barcodelecprice.value='';
	viewfrm.barcodematprice.value='';
	viewfrm.itemid.value='';
	viewfrm.OrderSerial.value = OrderSerial;
	viewfrm.startdate.value = startdate;
	viewfrm.entryname.value = entryname;
	viewfrm.lec_title.value = lec_title;
	viewfrm.SubTotalPrice.value = SubTotalPrice;
	viewfrm.barcodelecprice.value = barcodelecprice;
	viewfrm.barcodematprice.value = barcodematprice;		
	viewfrm.itemid.value=itemid;
	viewfrm.action='/academy/lecture/inc_lecturer_search.asp';
	viewfrm.target='view';
	viewfrm.submit();		
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
<iframe id="view" name="view" src="" width=10 height=10></iframe>
<!-- 바코드 출력을 위한 폼-->
<form name="viewfrm" method="post">
	<input type="hidden" name="OrderSerial">
	<input type="hidden" name="startdate">
	<input type="hidden" name="entryname">
	<input type="hidden" name="lec_title">
	<input type="hidden" name="SubTotalPrice">
	<input type="hidden" name="barcodelecprice">
	<input type="hidden" name="barcodematprice">
	<input type="hidden" name="itemid">
</form>
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
    	<input type="radio" name="searchfield" value="orderserial" <% if searchfield="orderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.orderserial)"> 주문번호
		<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'orderserial'); FocusAndSelect(frm, frm.orderserial);">

		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> 아이디
		<input type="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> 구매자명
		<input type="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> 구매자핸드폰
		<input type="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">
		<br>
		<input type="radio" name="searchfield" value="itemid" <% if searchfield="itemid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.itemid)"> 강좌번호
		<input type="text" name="itemid" value="<%= itemid %>" size="10" maxlength="10" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'itemid'); FocusAndSelect(frm, frm.itemid);">
        <input type="text" name="lecOption"  value="<%= lecOption %>" size="4" maxlength="4" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'itemid'); FocusAndSelect(frm, frm.lecOption);">

        <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> 기타조건
		<select name="etcfield">
		  <option value="">선택</option>
              <!--
              <option value="01" <% if etcfield="01" then response.write "selected" %> >구매자 명</option>
              -->
              <option value="02" <% if etcfield="02" then response.write "selected" %> >수령인 명</option>
              <!--
              <option value="03" <% if etcfield="03" then response.write "selected" %> >아이디</option>
              -->
              <option value="04" <% if etcfield="04" then response.write "selected" %> >입금자 명</option>
              <option value="06" <% if etcfield="06" then response.write "selected" %> >결제금액</option>
              <!--
              <option value="07" <% if etcfield="07" then response.write "selected" %> >구매자 핸드폰</option>
              -->
              <option value="08" <% if etcfield="08" then response.write "selected" %> >수령인 핸드폰</option>
              <!--
              <option value="09" <% if etcfield="09" then response.write "selected" %> >송장번호</option>
            	-->
            </select>
		<input type="text" name="etcstring" value="<%= etcstring %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
		<br>
		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
		주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        <input type="checkbox" name="checkJumunDiv" value="Y" <% if checkJumunDiv="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
		주문구분 :
		<select name="jumundiv">
	  	<option value="">선택</option>
          <option value="8" <% if jumundiv="8" then response.write "selected" %> >강좌주문</option>
        </select>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
	
</form>
</table>
<!---- /검색 ---->
<br>
<% if (searchfield = "itemid") then %>
	<!-- 강좌 설명 -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF">
			<td width=120 bgcolor="#DDDDFF">강좌코드</td>
			<td width=120 ><%= itemid %></td>
			<td width=120 bgcolor="#DDDDFF">강좌월구분</td>
			<td width=120 ><b><%= olecture.FOneItem.Flec_date %></b> <% if (olecture.FOneItem.isWeClass) then %><b><font color=red>단체강좌</font></b><% end if %></td>
			<td width=300 colspan="2" rowspan="3" ><img src="<%= olecture.FOneItem.Fbasicimg %>" width="150"></td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">강좌명</td>
			<td ><%= olecture.FOneItem.Flec_title %></td>
			<td bgcolor="#DDDDFF">검색어</td>
			<td ><%= olecture.FOneItem.Fkeyword %></td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">브랜드</td>
			<td colspan="3"><%= olecture.FOneItem.Flecturer_id %> (<%= olecture.FOneItem.Flecturer_name %>)</td>
		</tr>
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr  bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">수강료/매입가</td>
			<td >
			<%= FormatNumber(olecture.FOneItem.Flec_cost,0) %> / <%= FormatNumber(olecture.FOneItem.Fbuying_cost,0) %>
			</td>
			<td bgcolor="#DDDDFF">재료비</td>
			<td bgcolor="#FFFFFF" >
			<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
			포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% else %>
			별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">마일리지</td>
			<td >
			<%= olecture.FOneItem.Fmileage %> (point)
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">마감여부</td>
			<td >
			<% if olecture.FOneItem.IsSoldOut then %>
			<font color="#CC3333"><b>마감</b></font>
			<% else %>
			접수중
			<% end if %>
			<br> (마감기준 : 접수마감, 접수기간이외, 신청인원 정원초과, 전시안함, 사용안함 )
			</td>
			<td bgcolor="#DDDDFF">접수여부</td>
			<td >
			<% if olecture.FOneItem.Freg_yn="Y" then %>
			접수중
			<% else %>
			<font color="#CC3333">접수마감</font>
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">접수기간</td>
			<td >
			<%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">정원-신청 <br>= 남은인원</td>
			<td bgcolor="#FFFFFF" >
			  <%= olecture.FOneItem.Flimit_count %> 명
			-
			  <%= olecture.FOneItem.Flimit_sold %> 명
			=
			  <%= olecture.FOneItem.GetRemainNo %> 명
			</td>
			<td bgcolor="#DDDDFF">최소인원</td>
			<td bgcolor="#FFFFFF" colspan="4">
			<%= olecture.FOneItem.Fmin_count %> 명
			</td>
		</tr>
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">강의횟수 및 시간</td>
			<td bgcolor="#FFFFFF">
				<%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
			</td>
			<td bgcolor="#DDDDFF" rowspan="<%= olecschedule.FResultCount  %>">강의시작일</td>
			<td bgcolor="#FFFFFF" colspan="2">
				<%= olecture.FOneItem.Flec_startday1 %> ~ <%= olecture.FOneItem.Flec_endday1 %>				
				<% if (olecture.FOneItem.Flec_startday1<>olecschedule.FItemList(0).Fstartdate) or (olecture.FOneItem.Flec_endday1<>olecschedule.FItemList(0).Fenddate) then %>
					<br><b><%= olecschedule.FItemList(0).Fstartdate %> ~ <%= olecschedule.FItemList(0).Fenddate %></b>
				<% end if %>
			</td>
			<td ><%= getWeekdayStr(Left(olecture.FOneItem.Flec_startday1,10)) %></td>
		</tr>
<!--
		<% for i=1 to olecschedule.FResultCount-1 %>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#FFFFFF" >
			<%= olecschedule.FItemList(i).Fstartdate %> ~ <%= olecschedule.FItemList(i).Fenddate %>
			</td>
			<td><%= getWeekdayStr(Left(olecschedule.FItemList(i).Fstartdate,10)) %></td>
		</tr>
		<% next %>
-->
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" >전시여부</td>
			<td >
			<% if olecture.FOneItem.Fdisp_yn="Y" then %>
			전시
			<% else %>
			<font color="#CC3333">전시안함</font>
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF" >사용여부</td>
			<td colspan="3">
			<% if olecture.FOneItem.Fisusing="Y" then %>
			사용
			<% else %>
			<font color="#CC3333">사용안함</font>
			<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" >약도</td>
			<td >

			<% if olecture.FOneItem.Flec_mapimg<>"" then %>
				<a href="javascript:NewWindow('<%= olecture.FOneItem.Flec_mapimg %>');"><img src="http://www.thefingers.co.kr/images/d_2.gif" width="62" height="17" border="0"></a>
			<% end if %>

			</td>
			<td bgcolor="#DDDDFF" >등록일</td>
			<td colspan="3">
			<%= olecture.FOneItem.Fregdate %>
			</td>
		</tr>

		<tr  bgcolor="#FFFFF">
			<td colspan="6"><span style="cursor:hand" onclick="javascript:window.open('/academy/lecture/lib/LecUserList.asp?itemid=<%= itemid %>','checkwin','width=500,height=650,resizable,status=no,scrollbars=auto')">강좌 출석부 출력</span>&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onclick="ExcelPrint()">출석부 엑셀 저장</span></td>
		</tr>
	</table>
	<br>
	<!-- 리스트 시작 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oLectOption.FResultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oLectOption.FResultCount %></b>						
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>옵션코드</td>
    	<td>옵션명</td>
    	<td>접수기간</td>
    	<td>강좌일</td>
    	<td>남은인원</td>
    	<td>대기인원</td>
    	<td>마감여부</td>
	</tr>
	<% for i=0 to oLectOption.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
    	<td <%= chkIIF(oLectOption.FItemList(i).FlecOption=lecOption,"bgcolor=#DDDDDD","") %> ><a href="?searchfield=<%= searchfield %>&itemid=<%=oLectOption.FRectidx %>&lecOption=<%=oLectOption.FItemList(i).FlecOption%>&menupos=<%=menupos%>"><%=oLectOption.FItemList(i).FlecOption%></a></td>
    	<td><%=oLectOption.FItemList(i).FlecOptionName%></td>
    	<td><%=FormatDateTime(oLectOption.FItemList(i).FRegStartDate,2) & "~" & FormatDateTime(oLectOption.FItemList(i).FRegEndDate,2)%></td>
    	<td><%=FormatDateTime(oLectOption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oLectOption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oLectOption.FItemList(i).FlecEndDate,4)%></td>
    	<td><%=oLectOption.FItemList(i).Flimit_count & "명-" & oLectOption.FItemList(i).Flimit_sold & "명= " & (oLectOption.FItemList(i).Flimit_count-oLectOption.FItemList(i).Flimit_sold) & "명"%></td>
    	<td><%=oLectOption.FItemList(i).Fwait_count%>명</td>
    	<td><% if oLectOption.FItemList(i).IsOptionSoldOut then Response.Write "마감"%></td>
	</tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
	</table>
	<br>
	
	<!-- 리스트 시작 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ojumun.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="30">구분</td>
    	<td width="70">주문번호</td>
    	<td width="50">거래상태</td>
    	<td width="60">결제방법</td>
    	<td width="60">총 결제금액</td>
    	<td width="90">UserID</td>
    	<td width="40">수량</td>
    	<td width="80">옵션</td>
    	<td width="60">구매자 성명</td>
    	<td width="60">수강생</td>
    	<td width="60">수강생Hp</td>
    	<td width="70">주문일</td>
    	<td width="70">입금일</td>
    	<td>비고</td>
    </tr>
	<% for ix=0 to ojumun.FresultCount-1 %>
	
	<% if ojumun.FItemList(ix).IsAvailJumun then %>
	<% totalavailcount = totalavailcount + ojumun.FItemList(ix).FItemNo %>	
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
	<% else %>
	<tr align="center" bgcolor="#eeeeee" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='eeeeee';>
	<% end if %>
		<td><font color="<%= ojumun.FItemList(ix).CancelStateColor %>"><%= ojumun.FItemList(ix).CancelStateStr %></font></td>
		<td><a href="javascript:GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
		<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
		<td align="left"><a href="?searchfield=userid&userid=<%= ojumun.FItemList(ix).FUserID %>"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></a></font></td>
		<td><%= ojumun.FItemList(ix).FItemNo %></td>
		<td><%= ojumun.FItemList(ix).FItemoptionName %></td>
		<td><%= ojumun.FItemList(ix).FBuyName %></td>
		<td><%= ojumun.FItemList(ix).Fentryname %></td>
		<td><%= ojumun.FItemList(ix).Fentryhp %></td>
		<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		<td>
		    <% IF (ojumun.FItemList(ix).isWeClass) and (Not ojumun.FItemList(ix).isWeClassFixedOrder) then %>
		    <!-- input type="button" onClick="" value="단체강좌 관리" 주문번호 클릭 후 수정-->
		    <% else %>
			<input type="button" onClick="SongJangPrintProc_Fingers('<%= ojumun.FItemList(ix).FOrderSerial %>','<%= FormatDate(ojumun.FItemList(ix).flecStartDate,"0000/00/00") %>','<%= ojumun.FItemList(ix).Fentryname %>','<%= olecture.FOneItem.Flec_title %>','<%= ojumun.FItemList(ix).barcodesumprice %>','<%= formatnumber(ojumun.FItemList(ix).barcodelecprice,0)%>','<%=ojumun.FItemList(ix).barcodematprice%>','<%=ojumun.FItemList(ix).fitemid%>');" class="button" value="출력">
			<% end if %>
		</td>
    </tr>   	
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6"></td>
		<td align="center"><%= totalavailcount %></td>
		<td colspan="7"></td>
	</tr>	
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	        <% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for ix=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
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
	</tr>
	</table>	
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="30">구분</td>
    	<td width="50">주문구분</td>
    	<td width="70">주문번호</td>
    	<td width="50">Site</td>
    	<td width="90">UserID</td>
    	<td width="60">구매자</td>
    	<td width="40">인원</td>
    	<td width="50">구매총액</td>
    	<td width="50">쿠폰</td>
    	<td width="50">마일리지</td>
    	<td width="50">SKT</td>
    	<td width="60">결제금액</td>
    	<td width="60">결제방법</td>
    	<td width="50">거래상태</td>
    	<td width="70">주문일</td>
    	<td width="70">입금일</td>
    </tr>
    <% if ojumun.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
    </tr>
    <% else %>

	<% for ix=0 to ojumun.FresultCount-1 %>

	<% if ojumun.FItemList(ix).IsAvailJumun then %>
	<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); " style="cursor:hand">
	<% else %>
	<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); " style="cursor:hand">
	<% end if %>
		<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
		<td><%= ojumun.FItemList(ix).GetJumunDivName %></td>
		<td><a href="javascript:GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
		<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
		<td align="left"><a href="?searchfield=userid&userid=<%= ojumun.FItemList(ix).FUserID %>"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></a></font></td>
		<td><%= ojumun.FItemList(ix).FBuyName %></td>
		<td><%= ojumun.FItemList(ix).Ftotalitemno %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Ftencardspend,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fmiletotalprice,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fspendmembership,0) %></td>
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>

		<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
		<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><%= ojumun.FItemList(ix).IpkumDivName %></acronym></font></td>
		<% else %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
		<% end if %>
		<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
	<!--
		<td><acronym title="<%= ojumun.FItemList(ix).Fbeadaldate %>"><%= Left(ojumun.FItemList(ix).Fbeadaldate,10) %></acronym></td>
	-->
	</tr>
	<% next %>
	</table>
	<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="FFFFFF">       
	    <td>
	        <% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for ix=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
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
	</tr>
    
	</table>	
	<% end if %>
<% end if %>
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name="xlfrm" method="post" action="">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="lecOption" value="<%= lecOption %>">
<input type="hidden" name="searchfield" value="itemid">
</form>
<script type="text/javascript">
<!--
function ExcelPrint() {
	xlfrm.target="iiframeXL";
	xlfrm.action="/lectureadmin/lecture/dolectrollbookexcel.asp";
	xlfrm.submit();
}
//-->
</script>
<%
set olecture = Nothing
set olecschedule = Nothing
set ojumun = Nothing
%>

<script language='javascript'>
	ChangeFormBgColor(frm);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->