<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : giftday_item.asp
' Discription : ����� giftday_item
' History : 2014.03.31 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , maintitle , subtitle 
	idx = requestCheckvar(request("idx"),16)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

Dim oSubItemList
set oSubItemList = new Cgiftday_list
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = idx
	If idx <> "" then
		oSubItemList.GetContentsItemList()
	End If 
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;
	
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/sitemaster/gift/day/giftday.asp";
	}
		
	$(function(){
		//������ư
		$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

		$( "#subList" ).sortable({
			placeholder: "ui-state-highlight",
			start: function(event, ui) {
				ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
			},
			stop: function(){
				var i=99999;
				$(this).parent().find("input[name^='sort']").each(function(){
					if(i>$(this).val()) i=$(this).val()
				});
				if(i<=0) i=1;
				$(this).parent().find("input[name^='sort']").each(function(){
					$(this).val(i);
					i++;
				});
			}
		});
	});

//����
function popSubEdit(subidx) {
<% if idx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�˻� �ϰ� ���
function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/gift/day/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayItem() {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}
</script>
<%
	If idx <> "" then
%>
<p><b>�� MDPICK ����</b></p>
<!-- // ��ϵ� ���� ��� --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oSubItemList.FTotalCount%> �� /
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
		    </td>
		    <td align="right">
		    	<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="60" />
<col span="3" width="0*" />
<col width="110" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>ǥ�ü���</td>
    <td>��뿩��</td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "[" & oSubItemList.FItemList(lp).FItemid & "]" & oSubItemList.FItemList(lp).Fitemname
    	end if
    %>
    </td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
</tr>
<% Next %>
</tbody>
</table>
</form>
<%
	End If 
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->