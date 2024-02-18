<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 개인정보 문서 파기 관리
' History : 2019.08.13 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/isms/personaldata_cls.asp"-->
<%
dim page, i, iPageSize, userid, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, downFileDelYN, downFileconfirmYN
	page = RequestCheckVar(getnumeric(request("page")),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	downFileDelYN = requestCheckVar(request("downFileDelYN"),1)
	downFileconfirmYN = requestCheckVar(request("downFileconfirmYN"),1)

if page="" then page=1
iPageSize = 50
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-31)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)
yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

userid = session("ssBctId")

dim odata
set odata  = new Cpersonaldata
	odata.FPageSize = iPageSize
	odata.FCurrPage = page
	odata.frectlogtype = "A"
	odata.frectdownFileGubun = "EXCEL"
	odata.FRectStartdate = fromDate
	odata.FRectEnddate = toDate
	odata.FRectqryuserid = userid
	odata.FRectdownFileDelYN = downFileDelYN
	odata.FRectdownFileconfirmYN = downFileconfirmYN
	odata.GetpersonaldataList
%>

<script type="text/javascript">

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

// 전체선택
function totalCheck(){
	var f = document.frmArr;
	var objStr = "idx";
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

function downFileDelArr(){
    //선택없을시 체크
    var ret = 0;
    for (i=0; i< document.getElementsByName("idx").length; i++){
        if (document.getElementsByName("idx")[i].checked == true){
            ret = ret + 1;
        }
    }
    if (ret == 0){
        alert("선택값이 없습니다.");
        return;
    }

    //입력체크
    for (i=0; i< frmArr.idx.length; i++){
        if (frmArr.idx[i].checked == true){
            if (frmArr.downFileDelYN[i].value=='Y'){
                alert('이미 파기된 문서가 선택되어 있습니다.');
                frmArr.downFileDelYN[i].focus();
                return;
            }
        }
    }

	var ret = confirm('개인정보 문서를 파기 하시겠습니까?');
	if (ret){
		frmArr.mode.value = "downFileDelArr";
		frmArr.target="_self"
		frmArr.action="/admin/isms/personaldata_process.asp";
		frmArr.submit();
	}
}

function downFileconfirmArr(){
    //선택없을시 체크
    var ret = 0;
    for (i=0; i< document.getElementsByName("idx").length; i++){
        if (document.getElementsByName("idx")[i].checked == true){
            ret = ret + 1;
        }
    }
    if (ret == 0){
        alert("선택값이 없습니다.");
        return;
    }

    //입력체크
    for (i=0; i< frmArr.idx.length; i++){
        if (frmArr.idx[i].checked == true){
            if (frmArr.downFileDelYN[i].value=='N'){
                alert('파기 이전 문서가 선택되어 있습니다.');
                frmArr.downFileDelYN[i].focus();
                return;
            }
            if (frmArr.downFileconfirmYN[i].value=='Y'){
                alert('이미 확인서가 작성된 문서가 선택되어 있습니다.');
                frmArr.downFileconfirmYN[i].focus();
                return;
            }
        }
    }

	var ret = confirm('확인서를 작성 하시겠습니까?');
	if (ret){
		window.open('','downFileconfirm','width=1280,height=960,scrollbars=yes,resizable=yes');
		frmArr.mode.value = "downFileconfirmArr";
		frmArr.target='downFileconfirm';
		frmArr.action="/admin/isms/personaldata_downFileconfirm.asp";
		frmArr.submit();
		downFileconfirm.focus();
	}
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
        * 기간 : <% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
    </td>	
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="frm.submit();">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        * 문서파기여부 : <% drawSelectBoxisusingYN "downFileDelYN", downFileDelYN, "" %>
        &nbsp;
        * 확인서작성여부 : <% drawSelectBoxisusingYN "downFileconfirmYN", downFileconfirmYN, "" %>
    </td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
        문서파기는 텐바이텐 어드민 상에서 개인정보 파일을 다운받으신 경우에만 하시면 됩니다.
        <br>
        개인정보 파일을 다운로드 받지 않은 직원은 검색되는 정보가 없게 되며,
        파일이 존재하는 경우에만 파기해 주시면 됩니다. 
        <br><br>
        <font color="red">
        1단계 : 다운로드 하신 항목을 선택하신후에 "문서파기" 버튼을 눌러 파기.
        <br>
        2단계 : 파기후 해당 항목을 "확인서작성" 버튼을 눌러 확인서를 작성해 주시면 됩니다.
        </font>
    </td>
    <td align="right">	
        <input type="button" class="button" value="문서파기" onclick="downFileDelArr()">
        <input type="button" class="button" value="확인서작성" onclick="downFileconfirmArr()">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<form action="post" name="frmArr" method="post" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="20">
        검색결과 : <b><%= odata.FTotalCount %></b>
        &nbsp;
        페이지 : <b><%= page %>/ <%= odata.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="ckall" onClick="totalCheck()"></td>
    <td>직원ID</td>
    <td>접속IP</td>	
    <td>다운로드일시</td>	
    <td>매뉴명</td>
    <td>파일구분</td>
    <td>문서<br>파기여부</td>	
    <td>문서파기날짜</td>
    <td>확인서<br>작성여부</td>
    <td>확인서작성날짜</td>
</tr>
<% if odata.FresultCount>0 then %>
    <% for i=0 to odata.FresultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td>
            <input type="checkbox" name="idx" value="<%= odata.FItemlist(i).fidx %>" onClick="AnCheckClick(this);" <% if odata.FItemlist(i).fdownFileconfirmYN="Y" then response.write " disabled" %>>
            <input type="hidden" name="downFileDelYN" value="<%= odata.FItemlist(i).fdownFileDelYN %>">
            <input type="hidden" name="downFileconfirmYN" value="<%= odata.FItemlist(i).fdownFileconfirmYN %>">
        </td>
        <td><%= odata.FItemlist(i).fqryuserid %></td>
        <td><%= odata.FItemlist(i).frefip %></td>
        <td><%= odata.FItemlist(i).fregdate %></td>
        <td align="left"><%= odata.FItemlist(i).fmenuname %></td>
        <td><%= odata.FItemlist(i).FdownFileGubun %></td>
        <td><%= odata.FItemlist(i).fdownFileDelYN %></td>
        <td><%= odata.FItemlist(i).fdownFileDelDate %></td>
        <td><%= odata.FItemlist(i).fdownFileconfirmYN %></td>
        <td><%= odata.FItemlist(i).fdownFileconfirmDelDate %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="15" align="center">
            <% if odata.HasPreScroll then %>
                <span class="list_link"><a href="#" onclick="GotoPage('<%= odata.StartScrollPage-1 %>'); return false;">[pre]</a></span>
            <% else %>
            [pre]
            <% end if %>
            <% for i = 0 + odata.StartScrollPage to odata.StartScrollPage + odata.FScrollCount - 1 %>
                <% if (i > odata.FTotalpage) then Exit for %>
                <% if CStr(i) = CStr(odata.FCurrPage) then %>
                <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                <% else %>
                <a href="#" onclick="GotoPage('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
                <% end if %>
            <% next %>
            <% if odata.HasNextScroll then %>
                <span class="list_link"><a href="#" onclick="GotoPage('<%= i %>'); return false;">[next]</a></span>
            <% else %>
            [next]
            <% end if %>
        </td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="20" align="center" class="page_link">[파기할 개인정보 문서가 없습니다]</td>
    </tr>
<% end if %>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->