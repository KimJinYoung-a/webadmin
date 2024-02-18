<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcls.asp" -->

<%

dim searchType, searchString, upfilestate, upfiledate
dim page, pageSize
dim sitegubun
dim notinputonly

upfilestate     = RequestCheckVar(request("upfilestate"),32)
searchType      = RequestCheckVar(request("searchType"),32)
searchString    = RequestCheckVar(request("searchString"),32)
page            = RequestCheckVar(request("page"),10)
pageSize        = RequestCheckVar(request("pageSize"),10)
upfiledate      = RequestCheckVar(request("upfiledate"),32)
sitegubun      	= RequestCheckVar(request("sitegubun"),32)
notinputonly    = RequestCheckVar(request("notinputonly"),32)

if page="" then page=1
if pageSize="" then pageSize=300
if upfilestate="" then upfilestate="notupload"

dim OrefundList
set OrefundList = new CCSRefund
OrefundList.FCurrPage           = page
OrefundList.FPageSize           = pageSize
OrefundList.FRectReturnmethod   = "R007"
OrefundList.FRectSearchType     = searchType
OrefundList.FRectSearchString   = searchString
OrefundList.FRectUpfiledate     = upfiledate
OrefundList.FRectNotInputOnly   = notinputonly

if upfilestate="confirm" then
    OrefundList.FRectCurrstate      = "B005"
else
    OrefundList.FRectCurrstate      = "B001"
    OrefundList.FRectUploadState    = upfilestate
end if

if (sitegubun = "10x10") then
	OrefundList.GetRefundRequireList
elseif (sitegubun = "academy") then
	OrefundList.GetRefundRequireAcademyList
else
	'검색안함
end if


dim i
%>

<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript">
function regConfirmMsg(iid,fin){
    var popwin = window.open('/cscenter/action/pop_ConfirmMsg.asp?id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regConfirmMsgAcademy(iid,fin){
    var popwin = window.open('/cscenterv2/cs/pop_ConfirmMsg.asp?id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goPage(page){
    frm_search.page.value = page;
    frm_search.submit();
}

function reSearch(frm){
    frm.submit();
}
function CheckSubmit(frm){

    if (!CheckExists(frm)){
        alert('선택된 내역이 없습니다.');
        return;
    }

    if (confirm('선택내역을 이체파일로 작성하시겠습니까?')){
	    if(frm.ckidx.length > 1) {
			for(i = 0; i < frm.ckidx.length; i++) {
				if (frm.ckidx[i].checked == true) {
				/*
				    if (frm.arrrebankaccount.value == "") {
				    	frm.arrrebankaccount.value = eval("frm.rebankaccount" + frm.ckidx[i].value).value.replace(/,/g, "");
				    } else {
				    	frm.arrrebankaccount.value = frm.arrrebankaccount.value + ", " + eval("frm.rebankaccount" + frm.ckidx[i].value).value.replace(/,/g, "");
				    }
				*/
				}
			}
		}else{
		    if (frm.ckidx.checked == true) {
		    	//frm.arrrebankaccount.value = eval("frm.rebankaccount" + frm.ckidx.value).value.replace(/,/g, "");
		    }
		}

        frm.mode.value = "regfile";
        frm.submit();
    }
}

function CheckSubmitOLD(frm){
    if (!CheckExists(frm)){
        alert('선택된 내역이 없습니다.');
        return;
    }

    if (confirm('선택내역을 이체파일로 작성하시겠습니까?')){
        frm.mode.value = "regfileOLD";
        frm.submit();
    }
}

function RollBackFile(iid){
    var frm = document.frmSubmit;
    if (confirm('작성 이전으로 변경하시겠습니까?')){
        frm.asid.value = iid;
        frm.submit();
    }
}

function CheckExists(frm){
    if(frm.ckidx.length>1){
		for(i=0;i<frm.ckidx.length;i++){
			if (frm.ckidx[i].checked){
			    return true;
			}
		}
	}else{
	    return frm.ckidx.checked;
	}

	return false;
}

function switchCheckBox(){
    var form=document.frm_list;

	if(form.ckidx.length>1){
		for(i=0;i<form.ckidx.length;i++){
			if(form.switchCheck.checked){
			    if ((form.rebankname[i].value.length<1)||(form.rebankaccount[i].value<1)||(form.rebankownername[i].value.length<1)) continue;
				form.ckidx[i].checked=true;
			}else{
				form.ckidx[i].checked=false;
			}
			AnCheckClick(form.ckidx[i]);
		}
	}else{
		if(form.switchCheck.checked){
		    if ((form.rebankname.value.length<1)||(form.rebankaccount.value<1)||(form.rebankownername.value.length<1)) return;
			form.ckidx.checked=true;
		}else{
			form.ckidx.checked=false;
		}
		AnCheckClick(form.ckidx);
	}
}

function popUpFileByDate(frm, comp){
    if (frm.sitegubun.selectedIndex == 0) {
    	alert('사이트를 선택해 주세요.');
    	return;
    }

    if (comp.value.length<1){
        alert('작성일을 선택해 주세요.');
        comp.focus();
        return;
    }

    var popwin = window.open('poprefundfile.asp?sitegubun=' + frm.sitegubun.value + '&upfiledate=' + comp.value,'popUpFileByDate','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popUpFileAll(){
    var popwin = window.open('poprefundfile.asp?sitegubun=' + frm.sitegubun.value,'popUpFileByDate','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function researchByUpfiledate(comp){
    if (comp.value.length<1){
        alert('작성일을 선택해 주세요.');
        comp.focus();
        return;
    }

    frm_search.upfiledate.value = comp.value;
    frm_search.submit();
}

function selectCheckAll(){

}

function Cscenter_Action_List_Academy(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenterv2/cs/frame.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function Cscenter_Action_List_Lecture(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenterv2/cs_lecture/frame.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

// CS처리 등록/수정
function PopCSActionEdit_Academy(id,mode){
    PopCSActionCom_Academy(id,'',mode,'','');

}

// CS처리 등록/수정 공통
function PopCSActionCom_Academy(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs/pop_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function PopCSActionEdit_Lecture(id,mode){
    PopCSActionCom_Lecture(id,'',mode,'','');

}

// CS처리 등록/수정 공통
function PopCSActionCom_Lecture(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs_lecture/pop_lec_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

</script>


<!-- 검색 시작 -->
<form name="frm_search" method="GET" action="" onSubmit="return false" style="margin:0;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="upfiledate" value="">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			사이트 :
		    <select class="select" name="sitegubun">
			    <option value="" <% if sitegubun="" then response.write "selected" %> >------
			    <option value="10x10" <% if sitegubun="10x10" then response.write "selected" %> >텐바이텐
			    <option value="academy" <% if sitegubun="academy" then response.write "selected" %> >아카데미
		    </select>
			&nbsp;/&nbsp;
			이체파일 작성상태 :
		    <!--
		    <input type="radio" name="upfilestate" value="all" <% if upfilestate="all" then response.write "checked" %> >전체
		    -->
		    <input type="radio" name="upfilestate" value="notupload" <% if upfilestate="notupload" then response.write "checked" %> >환불요청
		    <input type="radio" name="upfilestate" value="uploaded" <% if upfilestate="uploaded" then response.write "checked" %> >환불작성중
		    <input type="radio" name="upfilestate" value="confirm" <% if upfilestate="confirm" then response.write "checked" %> >확인요청
		    &nbsp;
		    <select class="select" name="searchType">
			    <option value="orderserial" <% if searchType="orderserial" then response.write "selected" %> >주문번호
			    <option value="userid" <% if searchType="userid" then response.write "selected" %> >아이디
			    <option value="customername" <% if searchType="customername" then response.write "selected" %> >고객명
			    <option value="rebankownername" <% if searchType="rebankownername" then response.write "selected" %> >예금주
		    </select>
			<input type="text" class="text" name="searchString" size="16" value="<%= searchString %>">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="reSearch(frm_search)">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
            계좌정보 입력상태 :
            <select class="select" name="notinputonly">
                <option value=""></option>
                <option value="Y" <%= CHKIIF(notinputonly="Y", "selected", "") %>>미입력</option>
                <option value="N" <%= CHKIIF(notinputonly="N", "selected", "") %>>입력완료</option>
            </select>
			&nbsp;/&nbsp;
			출력갯수 :
			<select class="select" name="pageSize">
				<option value="100">100</option>
				<option value="300">300</option>
				<option value="500">500</option>
				<option value="1000">1000</option>
				<option value="2000">2000</option>
			</select>
			<script type="text/javascript">document.frm_search.pageSize.value='<%=pageSize%>';</script>
		</td>
	</tr>
	</table>
</form>

<p>

<!-- 액션 시작 -->
<form name="frmTmp" style="margin:0;">
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if (upfilestate="notupload") then %>
		    <input type="button" class="button" value="선택 내역 이체파일작성" onClick="CheckSubmit(frm_list)"  >

		    <input type="button" class="button" value="기존방식 이체파일작성" onClick="CheckSubmitOLD(frm_list)" disabled>
		    <% elseif (upfilestate="uploaded") then %>
		    <!-- 작성일별 목록 -->
		    <%

		    dim OrefundGroup
		    set OrefundGroup = new CCSRefund
		    OrefundGroup.FRectCurrstate      = "B001"
		    OrefundGroup.FRectReturnmethod   = "R007"
		    OrefundGroup.FRectSearchType     = searchType
		    OrefundGroup.FRectSearchString   = searchString
		    OrefundGroup.FRectUploadState    = upfilestate

			if (sitegubun = "10x10") then
				OrefundGroup.GetRefundRequireByFileDate
			elseif (sitegubun = "academy") then
				OrefundGroup.GetRefundRequireByFileDateAcademy
			else
				response.write "<font color=red>사이트</font>를 선택하세요"
			end if

		    %>
		    작성일 :
		    <select class="select" name="upfiledate">
		        <option value="">작성중파일 전체
		    	<% for i=0 to OrefundGroup.FResultCount -1 %>
		        <option value="<%= OrefundGroup.FItemList(i).Fupfiledate %>" <% if upfiledate=OrefundGroup.FItemList(i).Fupfiledate then response.write "selected" %> ><%= OrefundGroup.FItemList(i).Fupfiledate %> (<%= OrefundGroup.FItemList(i).FCount %>건)
		    	<% next %>
		    </select>
		    <input type="button" class="button" value="내역 보기" onclick="researchByUpfiledate(frmTmp.upfiledate)" onFocus="this.blur();">
		    <input type="button" class="button" value="파일 보기" onclick="popUpFileByDate(frm_search, frmTmp.upfiledate);" onFocus="this.blur();">

		    <!--
		    &nbsp;&nbsp;|&nbsp;&nbsp;
		    <input type="button" class="button" value="미처리 전체 파일 보기" onclick="popUpFileAll();" onFocus="this.blur();">
		    -->
		    <%
		    set OrefundGroup = Nothing
		    %>
		    <% else %>

		    <% end if %>
		</td>
		<td align="right">
		Total : <%= OrefundList.FTotalCount %>건
		&nbsp;
		</td>
	</tr>
	</table>
</form>
<!-- 액션 끝 -->

<p>



<form name="frm_list" method="post" action="refundlist_process.asp" style="margin:0;">
<input type="hidden" name="mode" value="regfile">
<input type="hidden" name="sitegubun" value="<%= sitegubun %>">
<input type="hidden" name="arrrebankaccount" value="">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="20">
		<% if (upfilestate="notupload") then %>
		<input type="checkbox" name="switchCheck" onClick="switchCheckBox()">
		<% else %>
	    <% end if %>
		</td>
		<td width="50">사이트</td>
		<td width="80">주문번호</td>
		<td width="80">CS IDX</td>
		<td width="75">고객ID</td>
		<td width="50">고객명</td>
		<td width="60">환불방식</td>
		<td width="60">환불요청액</td>
		<td width="45">은행</td>
		<td width="70">계좌</td>
		<td width="50">예금주</td>
		<td width="70">등록일</td>
		<td>CS처리상태</td>
		<td>IBK상태</td>
		<td>환불파일<br>작성일</td>
		<td width="100">Action</td>
	</tr>
	<% for i=0 to OrefundList.FResultCount -1 %>
		<%
		if (OrefundList.FItemList(i).Fencmethod = "TBT") then
		    ''사용 안함.
			OrefundList.FItemList(i).Frebankaccount = TBTDecrypt(OrefundList.FItemList(i).FencAccount)
	    elseif (OrefundList.FItemList(i).Fencmethod = "PH1") then
	        OrefundList.FItemList(i).Frebankaccount = OrefundList.FItemList(i).Fdecaccount
	    elseif (OrefundList.FItemList(i).Fencmethod = "AE2") then
	        OrefundList.FItemList(i).Frebankaccount = OrefundList.FItemList(i).Fdecaccount
		end if
		%>
	<input type="hidden" name="rebankname" value="<%= OrefundList.FItemList(i).Frebankname %>">
	<input type="hidden" name="rebankaccount" value="<%= LEN(OrefundList.FItemList(i).Frebankaccount) %>">
	<input type="hidden" name="rebankownername" value="<%= OrefundList.FItemList(i).Frebankownername %>">

	<tr bgcolor="#FFFFFF" align="center" >
	    <td>
	        <% if (upfilestate="notupload") then %>
	        <input type="checkbox" name="ckidx" value="<%= OrefundList.FItemList(i).Fasid %>" onClick="AnCheckClick(this)">
	        <input type="hidden" name="rebankaccount<%= OrefundList.FItemList(i).Fasid %>" value="<%= LEN(OrefundList.FItemList(i).Frebankaccount) %>">
	        <% else %>
	        <% end if %>
	    </td>
	    <td><%= OrefundList.FItemList(i).Fsitegubun %></td>
		<% if (sitegubun = "10x10") then %>
	    <td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= OrefundList.FItemList(i).FOrderSerial %>','','')"><%= OrefundList.FItemList(i).FOrderSerial %></a></td>
		<td><a href="javascript:Cscenter_Action_List('<%= OrefundList.FItemList(i).FOrderSerial %>','','A003')"><%= OrefundList.FItemList(i).Fasid %></a></td>
	    <% elseif (sitegubun = "academy") then %>
	    <td>
        	<!-- 땜방 : B/Y = 강좌/DIY -->
        	<% if (Left(OrefundList.FItemList(i).FOrderSerial, 1) = "B") then %>
        	<a href="javascript:Cscenter_Action_List_Lecture('<%= OrefundList.FItemList(i).FOrderSerial %>','','')"><%= OrefundList.FItemList(i).FOrderSerial %></a>
        	<% else %>
        	<a href="javascript:Cscenter_Action_List_Academy('<%= OrefundList.FItemList(i).FOrderSerial %>','','')"><%= OrefundList.FItemList(i).FOrderSerial %></a>
        	<% end if %>
	    </td>
		<td></td>
	    <% else %>
	    <td><%= OrefundList.FItemList(i).FOrderSerial %></td>
		<td></td>
	    <% end if %>
	    <td><%= printUserId(OrefundList.FItemList(i).FUserid,2,"**") %></td>
	    <td><%= OrefundList.FItemList(i).FCustomername %></td>
	    <td><%= OrefundList.FItemList(i).FreturnmethodName %></td>
	    <td align="right"><%= FormatNumber(OrefundList.FItemList(i).Frefundrequire,0) %></td>
	    <td><%= OrefundList.FItemList(i).Frebankname %></td>
	    <td align="left"><%= DispAcctStar(OrefundList.FItemList(i).Frebankaccount,3,8) %></td>
	    <td><%= OrefundList.FItemList(i).Frebankownername %></td>
	    <td><acronym title="<%= OrefundList.FItemList(i).Fregdate %>"><%= Left(OrefundList.FItemList(i).Fregdate,10) %></acronym></td>
	    <td>
	        <%= OrefundList.FItemList(i).getUpLoadStateName %>
	        <% if (OrefundList.FItemList(i).IsConfirmMsgFinished) then %>
	            <br><font color="#CC33CC"><acronym title="<%= OrefundList.FItemList(i).Fconfirmfinishmsg %>">(확인완료)</acronym></font>
	        <% end if %>
	    </td>
	    <td><%= OrefundList.FItemList(i).getIBKstateName %>
	    <% if (OrefundList.FItemList(i).FIBK_ERR_MSG<>"") then %>
        <br>(<%= OrefundList.FItemList(i).FIBK_ERR_MSG %>)
        <% end if %>
        </td>
	    <td><%= OrefundList.FItemList(i).Fupfiledate %></td>
	    <td>
			<% if (sitegubun = "10x10") then %>

		        <% if (upfilestate="notupload") then %>
		        <input class="button" type="button" value="확인요청" onclick="regConfirmMsg('<%= OrefundList.FItemList(i).Fasid %>','');" >
		        <input class="button" type="button" value="수정" onclick="PopCSActionEdit('<%= OrefundList.FItemList(i).Fasid %>','editrefundinfo');" >
		        <% else %>
	    	        <% if OrefundList.FItemList(i).IsRollBackValid then %>
	    	        <input class="button" type="button" value="작성이전변경" onclick="RollBackFile('<%= OrefundList.FItemList(i).Fasid %>');" >
	    	        <% end if %>
		        <% end if %>

		    <% elseif (sitegubun = "academy") then %>

		        <% if (upfilestate="notupload") then %>
		        	<input class="button" type="button" value="확인요청" onclick="regConfirmMsgAcademy('<%= OrefundList.FItemList(i).Fasid %>','');" >
		        	<!-- 땜방 : B/Y = 강좌/DIY -->
		        	<% if (Left(OrefundList.FItemList(i).FOrderSerial, 1) = "B") then %>
		        	<input class="button" type="button" value="수정" onclick="PopCSActionEdit_Lecture('<%= OrefundList.FItemList(i).Fasid %>','editrefundinfo');" >
		        	<% else %>
		        	<input class="button" type="button" value="수정" onclick="PopCSActionEdit_Academy('<%= OrefundList.FItemList(i).Fasid %>','editrefundinfo');" >
		        	<% end if %>
		        <% else %>
	    	        <% if OrefundList.FItemList(i).IsRollBackValid then %>
	    	        <input class="button" type="button" value="작성이전변경" onclick="RollBackFile('<%= OrefundList.FItemList(i).Fasid %>');" >
	    	        <% end if %>
		        <% end if %>

		    <% else %>
		    	&nbsp;
		    <% end if %>
	    </td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	    <td colspan="16" align="center">
	        <!-- 페이지 시작 -->
			<%
				if OrefundList.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & OrefundList.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + OrefundList.StartScrollPage to OrefundList.FScrollCount + OrefundList.StartScrollPage - 1

					if i>OrefundList.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if OrefundList.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
	    </td>
	</tr>
	</table>
</form>
<%
set OrefundList = Nothing
%>
<form name="frmSubmit" method="post" action="refundlist_process.asp">
<input type="hidden" name="mode" value="rollbackfile">
<input type="hidden" name="sitegubun" value="<%= sitegubun %>">
<input type="hidden" name="asid" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
