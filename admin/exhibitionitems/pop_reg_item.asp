<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ȹ�� ��ǰ ���� ��ǰ ��� �˾�
' History : 2018-11-07 ����ȭ
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
dim mastercode , detailcode , mode , idx , itemid , pickitem
dim oExhibition
mastercode = request("mastercode")
detailcode = request("detailcode")
idx = request("idx")
Mode = request("mode")

IF Mode = "" THEN Mode = "add"
if pickitem = "" then pickitem = 0

IF idx <> "" THEN
    set oExhibition = new ExhibitionCls
        oExhibition.Frectidx = idx
        oExhibition.getExhibitionItem()
        
		if mastercode = "" then 
        mastercode = oExhibition.FItem.Fmastercode
		end if 

		if detailcode = "" then 
        detailcode = oExhibition.FItem.Fdetailcode
		end if 
        itemid = oExhibition.FItem.Fitemid
        pickitem = oExhibition.FItem.Fpickitem
	set oExhibition = nothing	
End IF
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">
	// ����ǰ �߰� �˾�
	function findProd() {
			var popwin;
			popwin = window.open("/admin/Diary2009/pop_additemlist.asp", "popup_item", "width=900,height=600,scrollbars=yes,resizable=yes");
			popwin.focus();
	}

    function chgselectbox(v) {
        if (v != '' ){
            location.href = "?idx=<%=idx%>&mode=<%=mode%>&mastercode="+v;
        } else {
            location.href = "?idx=<%=idx%>&mode=<%=mode%>";
        }
    }

	function regitem() {
		var frm = document.frmreg;
		if (!frm.mastercode.value) {
			alert('������ ���� ���ּ���.');
			return;
		}

		if (!frm.iid.value) {
			alert('��ǰ�ڵ带 ���� ���ּ���.');
			frm.iid.focus();
			return;
		}
		frm.submit();
	}

	document.domain = "10x10.co.kr";
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frmreg" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp">
		<input type="hidden" name="mode" value="<%= Mode %>">
		<input type="hidden" name="eidx" value="<%= idx %>">
		<table class="tbType1 listTb">
			<tr>
				<td>
					<table class="tbType1 listTb">
						<tr bgcolor="#FFFFFF" height="25">
							<td colspan="2" ><b>���� ��ǰ ���</b></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> ����</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<% DrawMainPosCodeCombo "mastercode", mastercode ,"onchange='chgselectbox(this.value);'" %>
                                <% if mastercode > 0 then %>
                                    <% DrawDetailSelectBox "detailcode" , detailcode , mastercode %>
                                <% end if %>
							</td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> ��ǰ�ڵ�</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" class="text" name="iid" id="iid" value="<%=ItemID%>">
								<input type="button" class="button" value="��ǰã��" onClick="findProd();">
							</td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> BEST Pick ����</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="pickitem" value="1" <%=chkiif(pickitem="1","checked","") %> id="usey"><label for="usey">Pick ����</label>
								<input type="radio" name="pickitem" value="0" <%=chkiif(pickitem="0","checked","") %> id="usen"><label for="usen">Pick ��������</label>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="regitem();" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<!-- ����Ʈ �� -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->