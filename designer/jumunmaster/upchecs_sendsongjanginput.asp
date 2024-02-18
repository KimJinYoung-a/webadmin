<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->

<%
'' �ù�� �ϰ�����
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onChange="AssignDeliverSelect(this);">
     <option value='' <%if selectedId="" then response.write " selected"%>>�ù�缱��</option><%
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

''����� �����ΰ��
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


''�⺻�ù��.
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
        alert("�ù�縦 �������ּ���!");
        frm.songjangdiv.focus();
   }else if(frm.songjangno.value == ""){
        alert("�����ȣ�� �������ּ���!");
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
	    alert(" - ���� �Է����ּ���! ");
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
        alert("�ϰ� ����� �ֹ��� �����ϼ���.");
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
		alert("�ϰ� ����� �ֹ��� �����ϼ���.");
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
        alert("���� �ֹ��� �����ϴ�.");
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
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    		if (frm.detailidx[i].checked){
    			if (frm.songjangdiv[i].value.length<1){
    				alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
    				frm.songjangdiv[i].focus();
    				return;
    			}else if (trim(frm.songjangno[i].value).length<1){
    				alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
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
				alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
				return;
			}else if (trim(frm.songjangno.value).length<1){
				alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
				frm.songjangno.focus();
				return;
			}
		}
		orderserialArr = orderserialArr + frm.orderserial.value + ",";
		songjangnoArr  = songjangnoArr   + frm.songjangno.value + ",";
		songjangdivArr = songjangdivArr + frm.songjangdiv.value + ",";
		detailidxArr   = detailidxArr + frm.detailidx.value + ",";
    }


	if (confirm("���� �ֹ� �����͸� ��� �Ϸ� ó�� �Ͻðڽ��ϱ�?")){
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
		alert("������� ������ �����ϼ���.");
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
        alert('�˻� ������ �Է��ϼ���.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderserial")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('�˻� ������ ���ڸ� �����մϴ�.');
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


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" onsubmit="chksubmit(); return false">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" bgcolor="#FFFFFF">
			<select class="select" name="searchType" >
				<option value="">�˻�����</option>
				<option value="orderserial" <%= ChkIIF(searchType="orderserial","selected","") %> >�ֹ���ȣ</option>
				<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >��ǰ�ڵ�</option>
				<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >������</option>
				<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >������</option>
			</select>
			<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13" maxlength="11">
			&nbsp;
			�����:
			<select class="select" name="SearchGubun" OnChange="EnDisabledDateBox()">
				<option value="0" <% if SearchGubun="0" then response.write "selected" %> >����� ��ü
				<option value="1" <% if SearchGubun="1" then response.write "selected" %> >��� �Ϸ���
			</select>

			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			(�����)
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:chksubmit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
        	<input type="button" class="button" value="�����ֹ� ���ó��" onclick="CheckNFinish(document.frmbalju)">
		</td>
		<td align="right">

		    <input type="button" class="button" value="���� File �ϰ����" onclick="BatchSongjangInput(document.frmbalju)">
		   <!--
		    <input type="button" class="button" value="��ó�� ���� �����ȣ �ϰ����" onclick="BatchSongjangInputALL(frmbalju)">
		    -->
		</td>
	</tr>
</table>
<!-- �׼� �� -->

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
			�˻���� : <b><% = ojumun.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ojumun.FTotalpage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
		<td width="80">���ֹ���ȣ</td>
		<td width="50">�ֹ���</td>
		<td width="50">������</td>

		<td>��������</td>
		<td>����</td>
		<td>��������</td>

		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
		<td width="30">����</td>
		<td width="65">�����</td>
		<td width="65">�����</td>
		<td width="100"><% drawSelectBoxDeliverCompanyAssign "defaultsongjangdiv","" %></td>
		<td width="100" align="center">������ȣ</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center">[�˻������ �����ϴ�.]</td>
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