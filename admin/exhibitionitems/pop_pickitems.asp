<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
'###########################################################
' Description :  mdpick ���� ����
' History : 2018-11-06 ����ȭ����
'###########################################################
    Dim oExhibition, idx, page, i
    dim mastercode ,  detailcode
    mastercode =  requestCheckvar(request("mastercode"),10)
    detailcode =  requestCheckvar(request("detailcode"),10)

    SET oExhibition = new ExhibitionCls
        oExhibition.FrectMasterCode = mastercode
        oExhibition.FrectDetailCode = detailcode
        oExhibition.getExhibitionBestItemList
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
function jsSortIsusing() {
	var chk=0;
	$("#subList").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.mode.value = "mdpicksortingedit";
		document.frmList.action="/admin/exhibitionitems/lib/exhibition_proc.asp";
		document.frmList.submit();
	}
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

$(function(){
	//������ư
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// sortable
	$( "#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td>&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});
</script>
</head>
<body>
<div class="contSectFix scrl">
    <div class="contHead">
		<div class="locate"><h2>��ȹ����ǰ���� &gt; <strong>BEST PICK ��������(<%=chkiif(detailcode = 0,getMasterCodeName(mastercode),getMasterCodeName(mastercode) &"-"& getDetailCodeName(mastercode,detailcode))%>)</strong></h2></div>
	</div>

	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
                <p class="btn2 cBk1 ftLt"><a href="javascript:jsSortIsusing('');"><span class="eIcon"><em class="fIcon">��ǰ �Ӽ� ����</em></span></a></p>
            </div>
			<div class="ftRt">
				
			</div>
		</div>
		<div class="tPad15">
            <form name="frmList" id="frmList" method="post" action="" style="margin:0px;">
            <input type="hidden" name="mastercode" value="<%=mastercode%>" />
            <input type="hidden" name="detailcode" value="<%=detailcode%>" />
            <input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div><input type="checkbox" name="chkA" onClick="chkAllItem();"></div></th>
					<th><div>��ǰ��ȣ</div></th>
                    <th><div>�̹���</div></th>
                    <th><div>��ǰ��</div></th>                    
					<th><div>�������</div></th>
					<% if detailcode = 0 then  %>
					<th><div>MDPICK ��뿩��</div></th>
					<th><div>����</div></th>
					<% end if %>
				</tr>
				</thead>
				<% If oExhibition.FTotalCount > 0 Then %>
                <tbody id="subList">
                    <% For i = 0 to oExhibition.FTotalCount -1 %>
                    <tr height="25" bgcolor="<%=chkiif(oExhibition.FItemList(i).FIsusing="o","FFFFFF","f1f1f1")%>" align="center">
                        <td><input type="checkbox" name="chkIdx" value="<%= oExhibition.FItemlist(i).Fidx %>"></td>
                        <td><%= oExhibition.FItemlist(i).Fitemid %></td>
                        <td><img src="<%= oExhibition.FItemlist(i).FImageList%>" width="50" height="50" ></td>
                        <td><%= oExhibition.FItemlist(i).fitemname %></td>                 
                        <td><input type="text" size="2" maxlength="2" name="sort<%=oExhibition.FItemlist(i).Fidx%>" value="<%=oExhibition.FItemlist(i).Fsorting%>" class="text"></td>
						<% if detailcode = 0 then %>
                        <td>
                            <span class="rdoUsing">
                            <input type="radio" name="isusing<%=oExhibition.FItemlist(i).Fidx%>" id="rdoUsing<%=i%>_1" value="1" <%=chkIIF(oExhibition.FItemlist(i).FIsusing="1","checked","")%> />
							<label for="rdoUsing<%=i%>_1">���</label>
							<input type="radio" name="isusing<%=oExhibition.FItemlist(i).Fidx%>" id="rdoUsing<%=i%>_2" value="0" <%=chkIIF(oExhibition.FItemlist(i).FIsusing="0","checked","")%> />
							<label for="rdoUsing<%=i%>_2">����</label>
                            </span>
                        </td>
						<td>
							<span class="rdoUsing">
								<input type="radio" name="optioncode<%=oExhibition.FItemlist(i).Fidx%>" id="optioncode<%=i%>_0" value="0" <%=chkiif(oExhibition.FItemlist(i).Foptioncode="" or oExhibition.FItemlist(i).Foptioncode=0 or isnull(oExhibition.FItemlist(i).Foptioncode),"checked","")%>/>
								<label for="optioncode<%=i%>_0">����</label>
								<input type="radio" name="optioncode<%=oExhibition.FItemlist(i).Fidx%>" id="optioncode<%=i%>_1" value="1" <%=chkiif(oExhibition.FItemlist(i).Foptioncode=1,"checked","")%>/>
								<label for="optioncode<%=i%>_1">������</label>
								<input type="radio" name="optioncode<%=oExhibition.FItemlist(i).Fidx%>" id="optioncode<%=i%>_2" value="2" <%=chkiif(oExhibition.FItemlist(i).Foptioncode=2,"checked","")%>/>
								<label for="optioncode<%=i%>_2">Ư��</label>
								<input type="radio" name="optioncode<%=oExhibition.FItemlist(i).Fidx%>" id="optioncode<%=i%>_3" value="3" <%=chkiif(oExhibition.FItemlist(i).Foptioncode=3,"checked","")%>/>
								<label for="optioncode<%=i%>_3">�ܵ�</label>
								<input type="radio" name="optioncode<%=oExhibition.FItemlist(i).Fidx%>" id="optioncode<%=i%>_4" value="4" <%=chkiif(oExhibition.FItemlist(i).Foptioncode=4,"checked","")%>/>
								<label for="optioncode<%=i%>_4">����Ʈ</label>
							</span>
						</td>
						<% else %>
						<input type="hidden" name="pickitem<%=oExhibition.FItemlist(i).Fidx%>" value="<%=oExhibition.FItemlist(i).FIsusing%>">
						<% end if %>
                    </tr>
                    <% Next %>
                <% else %>
                    <tr bgcolor="#FFFFFF">
                        <td colspan="<%=chkiif(detailcode=0,"7","6")%>">[�˻������ �����ϴ�.]</td>
                    </tr>
                <% end if %>
                </tbody>
			</table>
            </form>
		</div>
	</div>
</div>
<%
    SET oExhibition = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->