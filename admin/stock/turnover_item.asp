<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [����]��ǰȸ���� 
' History : 2015.05.27 ���ʻ����� ��
'			2016.03.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid, cdl, cdm, cds, d, i, page, mstart, OnlySellyn, OnlyIsUsing, onlyOutItem, onlyOldItem, mwdiv, danjongyn, limityn
dim research, ChulgoNo, TurnOverPro, yyyy1, mm1, yyyy2, mm2, monthgubun, excBaseRegItem, dispCate
	dispCate = requestCheckvar(request("disp"),16)
	makerid = request("makerid")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	research = request("research")
	OnlySellyn = request("OnlySellyn")
	OnlyIsUsing = request("OnlyIsUsing")
	onlyOutItem = request("onlyOutItem")
	onlyOldItem = request("onlyOldItem")
	mwdiv       = request("mwdiv")
	danjongyn   = request("danjongyn")
	limityn     = request("limityn")
	ChulgoNo    = request("ChulgoNo")
	TurnOverPro = request("TurnOverPro")
	monthgubun = request("monthgubun")
	excBaseRegItem = request("excBaseRegItem")

''if (research="") and (OnlyIsUsing="") then OnlyIsUsing="Y"
if (research="") and (onlyOutItem="") then onlyOutItem="on"
if (research="") and (onlyOldItem="") then onlyOldItem="on"
if (research="") and (mwdiv="") then mwdiv="MW"
if (research="") and (danjongyn="") then danjongyn="SN"
if (research="") and (excBaseRegItem="") then excBaseRegItem="Y"

if (ChulgoNo="") then ChulgoNo="5"
if (TurnOverPro="") then TurnOverPro="0.5"

if (page = "") then
        page = 1
end if

if (yyyy1 = "") then
	d = CStr(dateadd("m" ,-1, now()))
	yyyy1 = Left(d,4)
	mm1 = Mid(d,6,2)

	yyyy2 = yyyy1
	mm2   = mm1
end if

dim olistforout
set olistforout = new CSummaryItemStock
	olistforout.FRectYYYYMM = yyyy1 + "-" + mm1
	olistforout.FRectEndDate = yyyy2 + "-" + mm2
	olistforout.FRectMakerid = makerid
	olistforout.FPageSize = 600
	olistforout.FCurrPage = page
	olistforout.FRectCD1 = cdl
	olistforout.FRectCD2 = cdm
	olistforout.FRectCD3 = cds
	olistforout.FRectOnlySellyn = OnlySellyn
	olistforout.FRectOnlyIsUsing = OnlyIsUsing
	olistforout.FRectOnlyOldItem = onlyOldItem
	olistforout.FRectOnlyOutItem = OnlyOutItem
	olistforout.FRectMwDiv = mwdiv
	olistforout.FRectDanjongyn =danjongyn
	olistforout.FRectLimityn =limityn
	olistforout.FRectChulgoNo   = ChulgoNo
	olistforout.FRectTurnOverPro = TurnOverPro
	olistforout.FRectExcBaseRegItem = excBaseRegItem
	olistforout.FRectMonthGubun = monthgubun
	olistforout.FRectDispCate		= dispCate

	if (makerid<>"") then
	    olistforout.GetItemListTurnOver
	else
	    olistforout.GetBrandListTurnOver
	end if
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function popDetailByBrand(imakerid){
    var strUrl = '/admin/stock/turnover_item.asp?menupos=982';

    strUrl = strUrl + '&makerid=' + imakerid;
    strUrl = strUrl + '&research=on';
    strUrl = strUrl + '&yyyy1=' + frm.yyyy1.value;
    strUrl = strUrl + '&mm1=' + frm.mm1.value;
    strUrl = strUrl + '&yyyy2=' + frm.yyyy2.value;
    strUrl = strUrl + '&mm2=' + frm.mm2.value;
    strUrl = strUrl + '&OnlySellyn=' + frm.OnlySellyn.value;
    strUrl = strUrl + '&OnlyIsUsing=' + frm.OnlyIsUsing.value;
    if (frm.onlyOutItem.checked){
        strUrl = strUrl + '&onlyOutItem=on';
    }else{
        strUrl = strUrl + '&onlyOutItem=';
    }
    if (frm.onlyOldItem.checked){
        strUrl = strUrl + '&onlyOldItem=on';
    }else{
        strUrl = strUrl + '&onlyOldItem=';
    }

    strUrl = strUrl + '&mwdiv=' + frm.mwdiv.value;
    strUrl = strUrl + '&danjongyn=' + frm.danjongyn.value;
    strUrl = strUrl + '&limityn=' + frm.limityn.value;
    strUrl = strUrl + '&cdl=' + frm.cdl.value;
    strUrl = strUrl + '&cdm=' + frm.cdm.value;
    strUrl = strUrl + '&cds=' + frm.cds.value;
    strUrl = strUrl + '&ChulgoNo=' + frm.ChulgoNo.value;
    strUrl = strUrl + '&TurnOverPro=' + frm.TurnOverPro.value;

    var popwin = window.open(strUrl,'popDetailByBrand','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function changecontent(){
	//dummy
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<script>
function SubmitForm()
{
        document.frm.page.value = 1;
        document.frm.submit();
}
function GotoPage(pg)
{
        document.frm.page.value = pg;
        document.frm.submit();
}
</script>

<script type="text/javascript">
// AJAX ���α׷�
var parentFrmName = "frm";
var xmlHttp;
var xmlDoc;
var xmlHttpMode, xmlHttpParam1, xmlHttpParam2, xmlHttpParam3;
var xmlHttpDefaultSet;
var xmlProcessId = 0;

function Trim(str){
 return str.replace(/\s/g,""); // \ -> �������� �Դϴ�.
}

function createXMLHttpRequest() {
        if (window.ActiveXObject) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        } else if (window.XMLHttpRequest) {
                xmlHttp = new XMLHttpRequest();
        }
}

function startRequest( mode,cdl,cdm,cds) {

		xmlHttpMode = mode;
		xmlHttpParam1 = cdl;
		xmlHttpParam2 = cdm;
		xmlHttpParam3 = cds;


		//alert('mode=' + mode + ',cdl=' + cdl + ',cdm=' + cdm + ',cds=' + cds);
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", "/common/module/normal_action_response.asp?mode=" + mode + "&param1=" + cdl + "&param2=" + cdm + "&param3=" + cds, true);
        xmlHttp.send(null);
}

function callback() {
	if(xmlHttp.readyState == 4) {
            if(xmlHttp.status == 200) {
                    // �������� ����Ÿ ��ȯ
                    // ��ü(TXT) : xmlHttp.responseText
                    if (window.ActiveXObject) {
                            // XML �� ��ȯ�Ѵ�.
                            // �ؽ�Ʈ �պκп��� "<" ���� ���ڵ��� �����Ѵ�.(���鹮�� ���ſ�,  �̷��� ���ϸ� ��ȯ�� �ȵȴ� --)
                            xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
                            var rawXML = xmlHttp.responseText;
                            var filteredML;

                            var index = 0;
                            for (var i = 0; i < rawXML.length; i++) {
                                    if (rawXML.charAt(i) == "<") {
                                            index = i;
                                            break;
                                    }
                            }

                            filteredML = rawXML.substring(index);
                            xmlDoc.loadXML(filteredML);
                    } else if (window.XMLHttpRequest) {
                            xmlDoc = xmlHttp.responseXML;
                    }

                    process();
            } else if (xmlHttp.status == 204){
                    // �����Ͱ� �������� ���� ���
                    alert("����Ÿ�� �������� �ʽ��ϴ�.(CODE : 200)");
            } else if (xmlHttp.status == 500){
                    // �����߻���
                    alert("����Ÿ ������ ������ �߻��Ͽ����ϴ�.(CODE : 500)");
            }
    }

}

// ���⸸ �����Ѵ�. �ش� ���������� ajax �� �̿��� ���� ����Ÿ�� �������� ǥ���Ѵ�.
function process() {
	var frm = eval("document." + parentFrmName);
	var buf;
	var length = xmlDoc.getElementsByTagName("value1").length;

	if (xmlHttpMode=="cdl"){
		frm.cdl.length = (length*1+1);

		for (i=0;i<length;i++){
			frm.cdl.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cdl.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.cdl.options[i + 1].selected = true;
			}
		}

		//����Ʈ��
		if (xmlHttpParam1!="") { startRequest('cdm',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="cdm"){
		frm.cdm.length = (length*1 + 1);
		frm.cds.length = 1;
		for (i=0;i<length;i++){
			frm.cdm.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cdm.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.cdm.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.cdm.length>0)) frm.cdm.options[0].selected = true;
		if ((xmlHttpParam3=="")&&(frm.cds.length>0)) frm.cds.options[0].selected = true;

		//����Ʈ��
		if (xmlHttpParam2!="") { startRequest('cds',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="cds"){
		frm.cds.length = (length*1 + 1);

		for (i=0;i<length;i++){
			frm.cds.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cds.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam3){
				frm.cds.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam3=="")&&(frm.cds.length>0)) frm.cds.options[0].selected = true;
	}
}

//��ǰ����Ʈ �ٿ�
function jsItemDown(){
	document.frm.target="hidifr";
	document.frm.action="itemlist_csv.asp";
	document.frm.submit();
}
</script>
<style>
p {margin:0; padding:0; border:0; font-size:100%;}
i, em, address {font-style:normal; font-weight:normal;}
.xls, .down {background-image:url(/images/partner/admin_element.png); background-repeat:no-repeat;}
.btn2 {display:inline-block; font-size:11px !important; letter-spacing:-0.025em; line-height:110%; border-left:1px solid #f0f0f0; border-top:1px solid #f0f0f0; border-right:1px solid #cdcdcd; border-bottom:1px solid #cdcdcd; background-color:#f2f2f2; background-image:-webkit-linear-gradient(#fff, #e1e1e1); background-image:-moz-linear-gradient(#fff, #e1e1e1); background-image:-ms-linear-gradient(#fff, #e1e1e1); background-image:linear-gradient(#fff, #e1e1e1); text-align:center; cursor:pointer;}
.btn2 a {display:block; font-size:11px !important; text-decoration:none !important;}
.btn2 span {display:block;}
.btn2 span em {display:block; padding-top:7px; padding-bottom:4px; text-align:center;}

.fIcon {padding-left:33px;}
.eIcon {padding-right:25px;}

.btn2 .xls {background-position:-125px -135px;}
.btn2 .down {background-position:right -231px;}
.cBk1, .cBk1 a {color:#000 !important;}
	</style>
<form name="frm" method="get" action="" onsubmit="return false;">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<% DrawYMYMBox yyyy1, mm1, yyyy2, mm2 %>�� ���� ����
		<br>
		�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<br>
        <input type="checkbox" name="onlyOutItem" value="on" <% if (onlyOutItem = "on") then response.write "checked" end if %>><%= yyyy2 %>��<%= mm2 %>�� ��������ǰ
        [<input type="text" class="text" name="ChulgoNo" value="<%= ChulgoNo %>" size="2" >���̸����&ȸ���� <input type="text" class="text" name="TurnOverPro" value="<%= TurnOverPro %>" size="3" >����]
		&nbsp;
		<input type="checkbox" name="excBaseRegItem" value="Y" <% if (excBaseRegItem = "Y") then response.write "checked" end if %>>�������ǰ�� ���رⰣ�� ��ϻ�ǰ ����
		&nbsp;
	    <input type="checkbox" name="onlyOldItem" value="on" <% if (onlyOldItem = "on") then response.write "checked" end if %>>�Ż�ǰ����(3������ ���)
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitForm()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�Ǹ�:<% drawSelectBoxSellYN "OnlySellyn", OnlySellyn %>
        &nbsp;
        ���:<% drawSelectBoxUsingYN "OnlyIsUsing", OnlyIsUsing %>
         &nbsp;
        ����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
        &nbsp;
        ����:<% drawSelectBoxLimitYN "limityn", limityn %>
     	&nbsp;
     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;
		������:
		<select class="select" name="monthgubun">
			<option value=""></option>
			<option value="2" <% if (monthgubun = "2") then %>selected<% end if %> >1����~3����</option>
			<option value="5" <% if (monthgubun = "5") then %>selected<% end if %> >4����~6����</option>
			<option value="11" <% if (monthgubun = "11") then %>selected<% end if %> >7����~12����</option>
			<option value="23" <% if (monthgubun = "23") then %>selected<% end if %> >1��~2��</option>
			<option value="24" <% if (monthgubun = "24") then %>selected<% end if %> >2���ʰ�</option>
		</select>
		<br>
		����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		&nbsp;
		����ī�װ� :
		<select class="select" name="cdl" onchange="startRequest('cdm',this.value,'','')"  ></select>
        <select class="select" name="cdm" onchange="startRequest('cds',eval(parentFrmName).cdl.value,this.value,'')"   ></select>
        <select class="select" name="cds"    ></select>
     </td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>
<script language='javascript'>
document.onload = getOnload();

function getOnload(){
	startRequest('cdl','<%= cdl %>','<%= cdm %>','<%= cds %>');
}
</script>
<div style="padding:0px 0 5px 0;">
* ������ �������, <font color="red">����ڻ�-����</font>���� ���ۼ� ��ư�� �����ø� ������ ���ɴϴ�.
&nbsp;&nbsp;&nbsp;<span class="btn2 cBk1" style="vertical-align:top;"><a href="javascript:jsItemDown();"><span class="eIcon down"><em class="fIcon xls">��ǰ���</em></span></a></span>
</div>
<% if (makerid<>"") then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= olistforout.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= olistforout.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2" width="40">��ǰ<br>�ڵ�</td>
		<td rowspan="2" width="50">�̹���</td>
		<td rowspan="2" width="120">�귣��ID</td>
		<td rowspan="2">��ǰ��(�ɼ�)</td>
		<td rowspan="2" width="30">�ŷ�<br>����</td>
		<td rowspan="2" width="55">�����<br>[<%= yyyy2 %>��<br><%= mm2 %>������]</td>
		<td colspan="5"><%= yyyy1 %>-<%= mm1 %>~<%= yyyy2 %>-<%= mm2 %>�� ȸ����</td>
	
		<td colspan="5">��ǰ�Ӽ�</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	
		<td width="30">ON<br>���</td>
		<td width="30">OFF<br>���</td>
		<td width="30">���<br>�հ�</td>
		<td width="30">����<br>���</td>
		<td width="40">ȸ����</td>
	
		<td width="30">�Ǹ�</td>
		<td width="30">���</td>
		<td width="30">����</td>
		<td width="50">����</td>
	</tr>
	<% if olistforout.FResultCount>0 then %>
	<% for i=0 to olistforout.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="javascript:PopItemSellEdit('<%= olistforout.FItemList(i).FItemID %>');"><%= olistforout.FItemList(i).FItemID %></a></td>
		<td><img src="<%= olistforout.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left"><%= olistforout.FItemList(i).Fmakerid %></td>
		<td align="left">
		    <a href="javascript:PopItemDetail('<%= olistforout.FItemList(i).Fitemid %>','<%= olistforout.FItemList(i).Fitemoption %>')"><%= olistforout.FItemList(i).Fitemname %></a>
		    <% if olistforout.FItemList(i).FitemoptionName <> "" then %>
		    <br>
		    <font color="blue">[<%= olistforout.FItemList(i).FitemoptionName %>]</font>
		    <% end if %>
		</td>
		<td><font color="<%= mwdivColor(olistforout.FItemList(i).Fmwdiv) %>"><%= mwdivName(olistforout.FItemList(i).Fmwdiv) %></font></td>
		<td><%= olistforout.FItemList(i).Faccumchulgo*(-1) %></td>
		<td><%= olistforout.FItemList(i).Fsellno*(-1) %></td>
		<td><%= olistforout.FItemList(i).Foffchulgono*(-1) %></td>
		<td><b><%= (olistforout.FItemList(i).Fsellno + olistforout.FItemList(i).Foffchulgono)*(-1) %></b></td>
		<td><%= olistforout.FItemList(i).Frealstock %></td>
		<td>
		    <% if olistforout.FItemList(i).Frealstock<>0 then %>
		        <%= CLng((olistforout.FItemList(i).Fsellno+olistforout.FItemList(i).Foffchulgono)*-1/olistforout.FItemList(i).Frealstock*100)/100 %>
		    <% end if %>
		</td>
	
		<td><font color="<%= ynColor(olistforout.FItemList(i).Fsellyn) %>"><%= olistforout.FItemList(i).Fsellyn %></font></td>
		<td><font color="<%= ynColor(olistforout.FItemList(i).Fisusing) %>"><%= olistforout.FItemList(i).Fisusing %></font></td>
		<td>
			<font color="<%= ynColor(olistforout.FItemList(i).Flimityn) %>"><%= olistforout.FItemList(i).Flimityn %>
			<% if (olistforout.FItemList(i).Flimityn = "Y") then %>
					<br>
					(<%= olistforout.FItemList(i).GetLimitStr %>)
			<% end if %>
			</font>
		</td>
		<td>
			<%= fncolor(olistforout.FItemList(i).Fdanjongyn,"dj") %>
		</td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if olistforout.HasPreScroll then %>
				<a href="javascript:GotoPage(<%= olistforout.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			
			<% for i=0 + olistforout.StartScrollPage to olistforout.FScrollCount + olistforout.StartScrollPage - 1 %>
			    <% if i>olistforout.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
			<% else %>
				<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
			<% end if %>
			<% next %>
			
			<% if olistforout.HasNextScroll then %>
				<a href="javascript:GotoPage(<%= i %>)">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
	</table>

<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= olistforout.FTotalCount %></b>
			&nbsp;
			(�ִ� <%= olistforout.FPageSize %>�� �귣�� ǥ�õ˴ϴ�.)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30">NO</td>
		<td width="150">�귣��ID</td>
		<td width="70">����ǰ��</td>
		<td width="70">On���</td>
		<td width="70">Off���</td>
		<td width="70">����հ�</td>
		<td width="70">�������</td>
		<td width="70">ȸ����</td>
		<td >&nbsp;</td>
	</tr>
	<% if olistforout.FResultCount>0 then %>
	<% for i=0 to olistforout.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><%= (((page - 1) * olistforout.FPageSize) + i + 1) %></td>
	    <td><%= olistforout.FItemList(i).Fmakerid %></td>
	    <td><%= formatNumber(olistforout.FItemList(i).Fcnt,0) %></td>
	    <td><%= formatNumber(olistforout.FItemList(i).Fsellno,0) %></td>
	    <td><%= formatNumber(olistforout.FItemList(i).Foffchulgono*(-1),0) %></td>
	    <td><b><%= formatNumber(olistforout.FItemList(i).Fsellno-olistforout.FItemList(i).Foffchulgono,0) %></b></td>
	    <td><b><%= formatNumber(olistforout.FItemList(i).Frealstock,0) %></b></td>
	    <td>
	    	<b>
	        <% if (olistforout.FItemList(i).Frealstock<>0) then %>
	            <%= CLng((olistforout.FItemList(i).Fsellno-olistforout.FItemList(i).Foffchulgono)/olistforout.FItemList(i).Frealstock*100)/100 %>
	        <% end if %>
	        </b>
	    </td>
	    <td align="left"><a href="javascript:popDetailByBrand('<%= olistforout.FItemList(i).Fmakerid %>');">��������&gt;&gt;</a></td>
	</tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
	</table>

<% end if %>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<%
set olistforout = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
