<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_contentsmenu_setting.asp
' Discription : I��(������) �̺�Ʈ ��Ƽ ������ �޴� ����
' History : 2019.02.07 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim cEvtCont, ix
Dim eCode, menuidx, GroupItemPriceView, GroupItemCheck, GroupItemType
dim menudiv, viewsort, isusing, ArrcMultiContentsMenu

eCode = requestCheckVar(Request("eC"),10)
menuidx = requestCheckVar(Request("menuidx"),10)

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
	cEvtCont.fnGetMultiContentsMenu
    GroupItemPriceView = cEvtCont.FGroupItemPriceView
    GroupItemCheck = cEvtCont.FGroupItemCheck
    GroupItemType = cEvtCont.FGroupItemType
    menudiv = cEvtCont.Fmenudiv
    viewsort = cEvtCont.Fviewsort
    isusing = cEvtCont.Fisusing
    ArrcMultiContentsMenu=cEvtCont.fnGetMultiContentsMenuList
    set cEvtCont = nothing
else
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    ArrcMultiContentsMenu=cEvtCont.fnGetMultiContentsMenuList
    set cEvtCont = nothing
end if

function GetMenuDivName(menudiv)
    if menudiv="1" then
        GetMenuDivName="�Ѹ� �̹���"
    elseif menudiv="2" then
        GetMenuDivName="����"
    elseif menudiv="3" then
        GetMenuDivName="�귣�� ���丮"
    elseif menudiv="4" then
        GetMenuDivName="��� ��ǰ����Ʈ"
    elseif menudiv="5" then
        GetMenuDivName="������"
    end if
end function
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
	frm.submit();
}

function fnCheckMenudiv(objval){
    if (objval == "4"){
        document.all.TrainInfo.style.display="";
    }
    else{
        document.all.TrainInfo.style.display="none";
    }
}

function jsDeleteContents(menuidx){
    if(menuidx != ""){
        document.frm.menuidx.value=menuidx;
        document.frm.submit();
    }
}
</script>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>
        <form name="frmEvt" method="post" style="margin:0px;" action="contentsmenu_process.asp">
        <input type="hidden" name="imod" value="MI">
        <input type="hidden" name="isusing" value="Y">
        <input type="hidden" name="evt_code" value="<%=eCode%>">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>��Ƽ ������ �޴� ����</B></td>
		   	</tr>
		   	<tr>
		   		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">������ ����<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   		    <select name="menudiv" onchange="fnCheckMenudiv(this.value);">
                       <option value="1"<% if menudiv="1" then response.write " selected" %>>�Ѹ� �̹���</option>
                       <option value="2"<% if menudiv="2" then response.write " selected" %>>����</option>
                       <option value="3"<% if menudiv="3" then response.write " selected" %>>�귣�� ���丮</option>
                       <option value="4"<% if menudiv="4" then response.write " selected" %>>��� ��ǰ����Ʈ</option>
                       <option value="5"<% if menudiv="5" then response.write " selected" %>>������</option>
                    </select> <input type="button" class="btn" value="�߰�" onclick="jsEvtSubmit(this.form);return false;">
		   		</td>
		   	</tr>
		   <tr id="TrainInfo" style="display:<% If menudiv<>"4" Then Response.write "none"%>">
		   		<td bgcolor="<%= adminColor("tabletop") %>">������ ���ø� ����</td>
		   		<td bgcolor="#FFFFFF">
		   			<table class="a">
                        <tr>
                            <td>��� ���� : <input type="radio" name="GroupItemType" value="T"<% if GroupItemType="T" then response.write " checked" %> onclick="fnViewTempType('T');TnPriceViewCheck('Y');">������ ���� <input type="radio" name="GroupItemType" value="B"<% if GroupItemType="B" or GroupItemType="" then response.write " checked" %> onclick="fnViewTempType('B')">�̹������ε�</td>
                        </tr>
                        <tr style="display:<% If GroupItemType="T" Then Response.write "none"%>" id="grouptemp5">
                            <td>����Ʈ ���� : <input type="radio" name="GroupItemCheck" value="I" onclick="TnPriceViewCheck('Y');fnViewTempType('T')"<% if GroupItemCheck="I" or GroupItemCheck="" then response.write " checked" %>>������ ����Ʈ <input type="radio" name="GroupItemCheck" value="T" onClick="TnPriceViewCheck('N');fnViewTempType('B')"<% if GroupItemCheck="T" then response.write " checked" %>>������ �̵� <input type="radio" name="GroupItemCheck" value="B" onClick="TnPriceViewCheck('N');fnViewTempType('M')"<% if GroupItemCheck="B" then response.write " checked" %>>�귣�� �̵�</td>
                        </tr>
                        <tr id="priceview" style="display:<% if GroupItemCheck="I" or GroupItemCheck="" then %><%else%>none<% end if %>">
                            <td>���ݳ��� ���� : <input type="radio" name="GroupItemPriceView" value="Y"<% if GroupItemPriceView="Y" or GroupItemPriceView="" then response.write " checked" %>>����ǥ�� <input type="radio" name="GroupItemPriceView" value="N"<% if GroupItemPriceView="N" then response.write " checked" %>>���� ��ǥ��</td>
                        </tr>
                    </table>
		   		</td>
		   	</tr>
		</table>
        </form>
	</td>
</tr>
<tr>
	<td>
        <% If isArray(ArrcMultiContentsMenu) Then %>
        <table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
                <td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ����</td>
                <td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
                <td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
		   	</tr>
            <% For ix = 0 To UBound(ArrcMultiContentsMenu,2) %>
            <tr bgcolor="#FFFFFF">
		   		<td align="center"><%=ArrcMultiContentsMenu(0,ix)%></td>
                <td align="left"><%=GetMenuDivName(ArrcMultiContentsMenu(1,ix))%></td>
                <td align="center"><%=ArrcMultiContentsMenu(2,ix)%></td>
                <td align="center"><input type="button" class="btn" value="����" onclick="jsDeleteContents(<%=ArrcMultiContentsMenu(0,ix)%>);return false;"></td>
		   	</tr>
            <% Next %>
        </table>
        <% End If %>
	</td>
</tr>
<tr>
	<td width="100%" align="right" >
		<a href="javascript:self.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
<form name="frm" method="post" style="margin:0px;" action="contentsmenu_process.asp">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="imod" value="MD">
<input type="hidden" name="menuidx">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->