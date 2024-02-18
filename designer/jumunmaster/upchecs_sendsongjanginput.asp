<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->

<%
'' 택배사 일괄적용
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onChange="AssignDeliverSelect(this);">
     <option value='' <%if selectedId="" then response.write " selected"%>>택배사선택</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate,BeasongCom
dim dateback, SearchGubun
dim SearchType, SearchValue

nowdate = Left(CStr(now()),10)

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)
SearchType  = requestCheckVar(request("SearchType"),16)
SearchValue = requestCheckVar(request("SearchValue"),16)
SearchGubun = requestCheckVar(request("SearchGubun"),16)

if SearchGubun = "" then SearchGubun = "0"

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)


dim page
dim ojumun

page    = requestCheckVar(request("page"),9)
if (page="") then page=1

set ojumun = new CCSJumunMaster

''출고일 기준인경우
if (SearchGubun = "1") then
	'ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegStart = DateSerial(yyyy1 , mm1 , dd1)
	ojumun.FRectRegEnd   = searchnextdate
end if

ojumun.FPageSize = 200
ojumun.FCurrPage = page
ojumun.FScrollCount = 10
ojumun.FRectSearchType  = SearchType
ojumun.FRectSearchValue = SearchValue
ojumun.FRectDesignerID = session("ssBctID")

ojumun.DesignerCS_BeasongList

dim ix,iy


''기본택배사.
dim idefaultSongjangDiv
idefaultSongjangDiv = CStr(fnGetUpcheDefaultSongjangDiv(session("ssBctID")))
%>
<script language='javascript'>
function AssignDeliverSelect(comp){
    var frm = comp.form;
	var selecidx = comp.selectedIndex;
	var frm;

    if (frm.detailidx.length>1){
    	for (var i=0;i<frm.songjangdiv.length;i++){
    	    frm.songjangdiv[i][selecidx].selected=true;
    	}
    }else{
        frm.songjangdiv[selecidx].selected=true;
    }
}

function ShowOrderInfo(frm,orderserial){
    var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}


function ViewItem(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
}

function songjangok(frm){
   if (frm.songjangdiv.value == ""){
        alert("택배사를 선택해주세요!");
        frm.songjangdiv.focus();
   }else if(frm.songjangno.value == ""){
        alert("송장번호를 선택해주세요!");
        frm.songjangno.focus();
   }else{
        frm.check.checked = true;
   }
}

function CheckThis(comp,i){
    var frm = comp.form;

	if (comp.value.length>5){
	    if (frm.songjangno.length>1){
	        frm.detailidx[i].checked=true;
	        AnCheckClick(frm.detailidx[i]);
        }else{
            frm.detailidx.checked=true;
            AnCheckClick(frm.detailidx);
        }
	}
}

function AnCheckColor(e){
	if (e.value != "")
		hL(e);
	else
		dL(e);
}
function hL(E){
	while (E.tagName!="TR")
	{
		E=E.parentElement;
	}

	E.className = "H";
}

function dL(E){
	while (E.tagName!="TR"){
		E=E.parentElement;
	}

	E.className = "";
}
function dodacheck(frm){
  var strPass = frm.songjangno.value;
  var strLength = strPass.length;

	if (frm.songjangno.value.indexOf ('-',0) != -1){
	    alert(" - 빼고 입력해주세요! ");
        var tst = frm.songjangno.value.substring(0, (strLength) - 1);
	    frm.songjangno.value = tst;
	}

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.detailidx.length>1){
		for(i=0;i<frm.detailidx.length;i++){
		    if (!frm.detailidx[i].disabled){
    			frm.detailidx[i].checked = comp.checked;
    			AnCheckClick(frm.detailidx[i]);
			}
		}
	}else{
	    if (!frm.detailidx.disabled){
    		frm.detailidx.checked = comp.checked;
    		AnCheckClick(frm.detailidx);
    	}
	}
}

function BatchSongjangInputALL(frm){
    var popwin = window.open('','BatchSongjangInput','width=600,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();

    document.frmArrInput.idxArr.value="";
    document.frmArrInput.iSall.value ="on";
	document.frmArrInput.target = "BatchSongjangInput";
	document.frmArrInput.submit();
}

function BatchSongjangInput(frm){
    var idxArr = '';
    if (!frm.detailidx){
        alert("일괄 등록할 주문을 선택하세요.");
		return;
    }

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){

    	    if (frm.detailidx[i].checked){
    	        idxArr = idxArr + frm.detailidx[i].value + ',';
    	    }
    	}
    }else{
        if (frm.detailidx.checked){
            idxArr = frm.detailidx.value;
        }
    }

	if (idxArr.length<1) {
		alert("일괄 등록할 주문을 선택하세요.");
		return;
	}



    var popwin = window.open('','BatchSongjangInput','width=600,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();

    document.frmArrInput.idxArr.value=idxArr;
    document.frmArrInput.iSall.value ="";
	document.frmArrInput.target = "BatchSongjangInput";
	document.frmArrInput.submit();


}


function CheckNFinish(frm){
	var pass = false;
    var orderserialArr = "";
    var songjangnoArr  = "";
    var songjangdivArr = "";
    var detailidxArr   = "";

    if (!frm.detailidx){
        alert("선택 주문이 없습니다.");
		return;
    }

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    	    pass = (pass||frm.detailidx[i].checked);
    	}
    }else{
        pass = frm.detailidx.checked;
    }

	if (!pass) {
		alert("선택 주문이 없습니다.");
		return;
	}

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    		if (frm.detailidx[i].checked){
    			if (frm.songjangdiv[i].value.length<1){
    				alert("택배사를 선택하시기 바랍니다.");
    				frm.songjangdiv[i].focus();
    				return;
    			}else if (trim(frm.songjangno[i].value).length<1){
    				alert("송장번호를 입력하시기 바랍니다.");
    				frm.songjangno[i].focus();
    				return;
    			}

    			orderserialArr = orderserialArr + frm.orderserial[i].value + ",";
				songjangnoArr  = songjangnoArr   + frm.songjangno[i].value + ",";
				songjangdivArr = songjangdivArr + frm.songjangdiv[i].value + ",";
				detailidxArr   = detailidxArr + frm.detailidx[i].value + ",";
    		}
    	}
    }else{
        if (frm.detailidx.checked){
			if (frm.songjangdiv.value.length<1){
				alert("택배사를 선택하시기 바랍니다.");
				return;
			}else if (trim(frm.songjangno.value).length<1){
				alert("송장번호를 입력하시기 바랍니다.");
				frm.songjangno.focus();
				return;
			}
		}
		orderserialArr = orderserialArr + frm.orderserial.value + ",";
		songjangnoArr  = songjangnoArr   + frm.songjangno.value + ",";
		songjangdivArr = songjangdivArr + frm.songjangdiv.value + ",";
		detailidxArr   = detailidxArr + frm.detailidx.value + ",";
    }


	if (confirm("선택 주문 데이터를 출고 완료 처리 하시겠습니까?")){
	    frm.orderserialArr.value = orderserialArr;
	    frm.songjangnoArr.value  = songjangnoArr;
        frm.songjangdivArr.value = songjangdivArr;
        frm.detailidxArr.value   = detailidxArr;
		frm.submit();
	}
}


function trim(theString){
   var resultString = theString;

   if (theString.indexOf(" ") == 0) {
        resultString = theString.substring(1, theString.length);
   }

   if (resultString.lastIndexOf(" ") == resultString.length) {
        resultString = resultString.substring(1,theString.length-1);
   }

   return resultString
}

function BaljuReprint1(){
    var frm = document.frmbalju;
	var pass = false;

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    	    pass = (pass||frm.detailidx[i].checked);
    	}
    }else{
        pass = frm.detailidx.checked;
    }

	if (!pass) {
		alert("재출력할 내역을 선택하세요.");
		return;
	}else{
	    var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
 		frm.action = "reselectbaljulist11.asp";
		frm.submit();
	}

}

function EnDisabledDateBox(){
	var bool = (frm.SearchGubun.value=="0");
	document.frm.yyyy1.disabled = bool;
	document.frm.yyyy2.disabled = bool;
	document.frm.mm1.disabled = bool;
	document.frm.mm2.disabled = bool;
	document.frm.dd1.disabled = bool;
	document.frm.dd2.disabled = bool;
}

function chksubmit(){
    var frm = document.frm;

    if ((frm.searchType.value.length>0)&&(frm.searchValue.value.length<1)){
        alert('검색 내용을 입력하세요.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderserial")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('검색 내용은 숫자만 가능합니다.');
            frm.searchValue.focus();
            return;
        }
    }

    frm.submit();
}

function popMisendInput1(iidx){
    var popwin = window.open('popMisendInput11.asp?idx=' + iidx,'popMisendInput','width=440,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ViewCSDetail(detailidx) {
    var popwin = window.open("/designer/jumunmaster/upchecsdetail.asp?idx=" + detailidx,"ViewCSDetail");
    popwin.focus();
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" onsubmit="chksubmit(); return false">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" bgcolor="#FFFFFF">
			<select class="select" name="searchType" >
				<option value="">검색조건</option>
				<option value="orderserial" <%= ChkIIF(searchType="orderserial","selected","") %> >주문번호</option>
				<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >상품코드</option>
				<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >구매자</option>
				<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >수령인</option>
			</select>
			<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13" maxlength="11">
			&nbsp;
			출고여부:
			<select class="select" name="SearchGubun" OnChange="EnDisabledDateBox()">
				<option value="0" <% if SearchGubun="0" then response.write "selected" %> >미출고 전체
				<option value="1" <% if SearchGubun="1" then response.write "selected" %> >출고 완료일
			</select>

			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			(출고일)
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:chksubmit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
        	<input type="button" class="button" value="선택주문 출고처리" onclick="CheckNFinish(document.frmbalju)">
		</td>
		<td align="right">

		    <input type="button" class="button" value="송장 File 일괄등록" onclick="BatchSongjangInput(document.frmbalju)">
		   <!--
		    <input type="button" class="button" value="미처리 내역 송장번호 일괄등록" onclick="BatchSongjangInputALL(frmbalju)">
		    -->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmbalju" method="post" action="upchebeasong_cs_Process.asp">
    <input type="hidden" name="mode" value="SongjangInput">
    <input type="hidden" name="orderserialArr" value="">
    <input type="hidden" name="songjangnoArr" value="">
    <input type="hidden" name="songjangdivArr" value="">
    <input type="hidden" name="detailidxArr" value="">
    <input type="hidden" name="isall" value="">

	<tr bgcolor="FFFFFF">
		<td height="25" colspan="15">
			검색결과 : <b><% = ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
		<td width="80">원주문번호</td>
		<td width="50">주문인</td>
		<td width="50">수령인</td>

		<td>접수구분</td>
		<td>제목</td>
		<td>접수사유</td>

		<td width="50">상품코드</td>
		<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
		<td width="30">수량</td>
		<td width="65">등록일</td>
		<td width="65">출고일</td>
		<td width="100"><% drawSelectBoxDeliverCompanyAssign "defaultsongjangdiv","" %></td>
		<td width="100" align="center">운송장번호</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>

	<% for ix=0 to ojumun.FresultCount-1 %>
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).Forderserial %>">
	<tr align="center" class="a" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="detailidx" value="<%= ojumun.FMasterItemList(ix).Fidx %>" onClick="AnCheckClick(this);"></td>
		<td height="25">
			<a href="javascript:ShowOrderInfo(frmshow,'<%= ojumun.FMasterItemList(ix).FOrgOrderSerial %>')"><%= ojumun.FMasterItemList(ix).FOrgOrderSerial %></a>
    		<% if (ojumun.FMasterItemList(ix).Forderserial <> ojumun.FMasterItemList(ix).Forgorderserial) then %>
    			+
    		<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>

		<td><%= ojumun.FMasterItemList(ix).Fdivcdname %></td>
		<td><a href="javascript:ViewCSDetail(<%= ojumun.FMasterItemList(ix).Fmasteridx %>)"><%= ojumun.FMasterItemList(ix).Ftitle %></a></td>
		<td><%= ojumun.FMasterItemList(ix).Fgubun01name %> >> <%= ojumun.FMasterItemList(ix).Fgubun02name %></td>

		<td><%= ojumun.FMasterItemList(ix).FItemid %></td>
		<td align="left">
			<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
			<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
			<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fregdate %>"><%= left(ojumun.FMasterItemList(ix).Fregdate,10) %></acronym></td>

		<td><acronym title="<%= ojumun.FMasterItemList(ix).Ffinishdate %>"><%= left(ojumun.FMasterItemList(ix).Ffinishdate,10) %></acronym></td>
		<td>
		    <% if (IsNULL(ojumun.FMasterItemList(ix).FSongjangdiv) or (ojumun.FMasterItemList(ix).FSongjangdiv=0)) then  %>
		        <% drawSelectBoxDeliverCompany "songjangdiv",idefaultSongjangDiv %>
		    <% else %>
		        <% drawSelectBoxDeliverCompany "songjangdiv",ojumun.FMasterItemList(ix).FSongjangdiv %>
		    <% end if %>
		</td>
		<td><input type="text" class="text" name="songjangno" size="16" value="<%= ojumun.FMasterItemList(ix).FSongjangno %>" onKeyup="CheckThis(this,'<%= ix %>');" maxlength=16></td>
	</tr>
	<% next %>
<% end if %>
    </form>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
	</tr>
</table>

<%
set ojumun = Nothing
%>

<form name="frmshow" method="post">
<input type="hidden" name="orderserial" value="">

</form>

<form name="frmArrInput" method="post" action="upchecs_pop_BatchSongjangInput.asp">
<input type="hidden" name="idxArr" value="">
<input type="hidden" name="iSall" value="">
</form>
<script language='javascript'>
    document.onload = EnDisabledDateBox();
</script>


<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->