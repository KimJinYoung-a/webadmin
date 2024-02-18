<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : PC���ΰ��� MD��
' History : ������ ����
'			2022.07.01 �ѿ�� ����(isms�������ġ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_event_rotationcls.asp"-->
<%

dim i
dim page, malltype
dim isusing, research
dim itemid, sdate, edate, realdate, lowestPrice
dim realdatereset, getdate

page = request("page")
isusing = request("isusing")
research = request("research")
itemid = request("itemid")
sdate = request("iSD")
edate = request("iED")
realdate = request("realdate")
realdatereset = request("realdatereset")
getdate = request("getdate")

lowestPrice = request("lowestPrice")
if (page = "") then
        page = "1"
end if

if research = "" and isusing="" then isusing="Y"

if realdatereset = "1" then 
	realdate = ""
else
	if realdate = "" then realdate = date()
end if 
if getdate="" then getdate=realdate
'==============================================================================
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate

mdchoicerotate.FCurrPage = CInt(page)
mdchoicerotate.FPageSize = 100
mdchoicerotate.FRectIsUsing = isusing
mdchoicerotate.FRectItemID = itemid
mdchoicerotate.FRectSDate = sdate
mdchoicerotate.FRectEDate = edate
mdchoicerotate.FRectrealdate = realdate
mdchoicerotate.FRectIsLowestPrice = lowestPrice
mdchoicerotate.list

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript">
function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

function frmChange()
{
	var vfm = document.vfrm;
	if(confirm("<%=realdate%>���� ���ü����� ���� ��Ͽ� ���̴� ���� �״�� ��� ����˴ϴ�.\n��ü ���� �Ͻðڽ��ϱ�?"))
	{
		vfm.action="doMainMdChoiceChange.asp";
		vfm.submit()
	}
	else
		return;
}

var chkUsing="<%=isusing%>";
function usingAllChange()
{
	if(chkUsing=="Y") { chkUsing = "N"; }
	else { chkUsing = "Y"; }

	for (var i=0;i<document.vfrm.isusing.length;i++){
		document.vfrm.isusing[i].value=chkUsing;
	}
}

function writeItem(idx) {
	if(idx==0) {
		var mode = "write";
	} else {
		var mode = "modify";
	}
	var mdcWrPop = window.open("main_md_recommend_flash_write.asp?mode="+mode+"&idx="+idx,"popMdcWr","width=1200,height=700,scrollbars=yes");
	mdcWrPop.focus();
}

function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=1200,height=700,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

function popupMainPreview(realdate){			
	var popwin; 		
	popwin = window.open("/admin/sitemaster/main_preview.asp?realdate="+realdate, "popup_main_preview", "width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popRegArrayItem() {
	var popwin;
    var popwin = window.open('main_md_recommend_flash_writes.asp?realdate=<%=realdate%>','popRegArray','width=1200,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popCurrentStock(itemid) {
	var popwin;
    var popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=<%= request("menupos") %>&itemid='+itemid,'popRegArray','width=1200,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function checkdate() {
	var frm = document.frm;
		frm.realdate.value = '';
}

function delItem(idx) {
	var vFrm = document.delfrm;
	if(confirm("�ش� ��ǰ�� ���� �˴ϴ�.\n���� �Ͻðڽ��ϱ�?"))
	{
		vFrm.idx.value = idx;
		vFrm.realdatereset.value = (document.frm.realdatereset.checked) ? 1 : 0;
		vFrm.action="doMainMdChoiceChange.asp";
		vFrm.submit();
	} else {
		return;
	}
}

$(function() {
	<% if realdate <> "" then %>
	$("#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="45" colspan="13" style="border:1px solid #F9BD01;">&nbsp;</td>');
			$(".etcInfo").hide();
		},
		stop: function(){
			var i=0;
			$(this).find("input[name^='disporder']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='disporder']").each(function(){
				$(this).val(i);
				i++;
			});
			$(".etcInfo").show();
		}
	});
	<% end if %>
})

function fnMobileMDPickCopy(){
	if(confirm("����� MD Pick�� ������ ������ PC���ο� �����մϴ�.\n�������ðڽ��ϱ�?\n\n�� (����) PC���ο� �ٷ� �ݿ��˴ϴ�.")) {
		if($("#getdate").val()!=""){
			$.ajax({
				type: "POST",
				url: "ajaxMDRecommendCopy.asp",
				data: "getdate="+$("#getdate").val(),
				cache: false,
				success: function(message) {
					if(message=="OK") {
						alert("���� �Ϸ�");
						window.location.reload();
					} else {
						alert("���翡 �����߽��ϴ�.");
					}
				},
				error: function(err) {
					alert(err.responseText);
				}
			});
		}
		else{
			alert("ī������ �������ּ���.");
		}
	}
}
</script>

<form name="refreshFrm" method="post" style="margin:0;">
</form>

<form name="delfrm" method="post" action="">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="realdate" value="<%=realdate%>">
	<input type="hidden" name="mode" value="del" />
	<input type="hidden" name="realdatereset" value="" />
	<input type="hidden" name="idx" />
</form>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� :
		<input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=9 class="text"> &nbsp;/
		����� : 
		<input id="iSD" name="iSD" value="<%=sdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=edate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		 &nbsp; &nbsp; &nbsp;������ ���� : 
		<select name="lowestPrice" class="select">
		<option value="" >��ü
		<option value="Y" <%=chkIIF(lowestPrice="Y","selected","")%> >���
		<option value="N" <%=chkIIF(lowestPrice="N","selected","")%> >������
		</select> &nbsp;
		<div style="float:right;vertical-align:middle;padding-right:20px;">
			�������� : 
			<input id="realdate" name="realdate" value="<%=realdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="realdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> &nbsp;
			[ ��ü�˻� : <input type="checkbox" name="realdatereset" value="1" <%=chkiif(realdatereset = "1" ,"checked","")%> onclick="checkdate()" style="vertical-align:middle"/> ]
		</div>
		<br>
	</td>
	<script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "iSD", trigger    : "iSD_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "iED", trigger    : "iED_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "realdate", trigger    : "realdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
				$("input:checkbox[name='realdatereset']").prop("checked", false);
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0;">
	<tr>
		<td align="right">
			<input id="getdate" name="getdate" value="<%=getdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="getdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_End = new Calendar({
					inputField : "getdate", trigger    : "getdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<input type="button" class="button" value="�����MDPICK��������" onClick="fnMobileMDPickCopy();">
			<% if realdate <> "" then %>
			<input type="button" class="button" value="���ü�������" onClick="frmChange()">
			&nbsp;
			<% end if %>
			<input type="button" class="button" value="������ǰ���" onClick="javascript:popRegArrayItem();">
			&nbsp;
			<input type="button" class="button" value="�űԵ��" onClick="javascript:writeItem(0);">
			&nbsp;
			<input type="button" class="button" value="�̸�����" onClick="popupMainPreview('<%=realdate%>')">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<form name="vfrm" method="POST" action="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sUsing" value="<%= isusing %>">
<input type="hidden" name="realdate" value="<%=realdate%>">
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		�˻���� : <b><%=mdchoicerotate.Ftotalcount%></b>
		&nbsp;
		������ : <b><%=page%> / <%=mdchoicerotate.Ftotalpage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center">��ǰ�ڵ�</td>
	<td align="center" class="etcInfo">�̹���</td>
	<td align="center" class="etcInfo">��ǰ����</td>
	<td align="center">��ǰ����</td>
	<td align="center">���ü���</td>
    <td align="center" class="etcInfo">������</td>
    <td align="center" class="etcInfo">������</td>
	<td align="center" class="etcInfo">�����</td>
	<td align="center" class="etcInfo">������ ����</td>	
	<td align="center" class="etcInfo">�Ǹſ���</td>
	<td align="center" class="etcInfo">�����</td>
	<td align="center" class="etcInfo">�����۾���</td>
	<td align="center" class="etcInfo">��ǰ ����</td>
</tr>
<tbody id="subList">
<% for i=0 to mdchoicerotate.FResultcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<input type="hidden" name="idx" value="<%= mdchoicerotate.FItemList(i).Fidx %>">
			<a href="javascript:writeItem(<%= mdchoicerotate.FItemList(i).Fidx %>);"><%= mdchoicerotate.FItemList(i).Flinkitemid %></a>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).FTentenImg <> "" then %>
				<img src="<%= mdchoicerotate.FItemList(i).FTentenImg %>" border=0 width="56">
			<% else %>				
				<img src="<%= mdchoicerotate.FItemList(i).Fphotoimg %>" border=0 width="56">					
			<% end if %>
			<br/><button type="button" onClick="editItemImage('<%= mdchoicerotate.FItemList(i).Flinkitemid %>')">����</button>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).FItemDiv = "21" then %>
				<span style="color:blue">�� ��ǰ</span>			
			<% else %>
				<span style="color:red">�Ϲ� ��ǰ</span>							
			<% end if %>			
		</td>				
		<td align="left">
			
			<a href="javascript:writeItem(<%= mdchoicerotate.FItemList(i).Fidx %>);">
				<%=chkIIF(mdchoicerotate.FItemList(i).Ftextinfo="" or isnull(mdchoicerotate.FItemList(i).Ftextinfo),"","TEXT : " & ReplaceBracket(mdchoicerotate.FItemList(i).Ftextinfo) & "<br>") %>
				LINK : <%= ReplaceBracket(mdchoicerotate.FItemList(i).Flinkinfo) %>
			</a>
			<% if mdchoicerotate.FItemList(i).Flinkitemid > 0 then %>
			<table cellpadding="1" cellspacing="1" class="a etcInfo" style="padding-top:15px;">
				<tr>
					<td>�ǸŰ� :</td>
					<td><%=FormatNumber(mdchoicerotate.FItemList(i).Forgprice,0)%> <%=mdchoicerotate.FItemList(i).saleCouponPriceCheck(mdchoicerotate.FItemList(i).Fsailyn , mdchoicerotate.FItemList(i).FitemCouponYn , mdchoicerotate.FItemList(i).Forgprice , mdchoicerotate.FItemList(i).Fsailprice , mdchoicerotate.FItemList(i).FitemCouponType)%></td>
				</tr>
				<tr>
					<td>���� :</td>
					<td><%=fnPercent(mdchoicerotate.FItemList(i).Forgsuplycash,mdchoicerotate.FItemList(i).Forgprice,1)%> <%=mdchoicerotate.FItemList(i).priceMarginCheck( mdchoicerotate.FItemList(i).Fsailyn , mdchoicerotate.FItemList(i).FitemCouponYn , mdchoicerotate.FItemList(i).FitemCouponType , mdchoicerotate.FItemList(i).Fsailsuplycash , mdchoicerotate.FItemList(i).Fsailprice , mdchoicerotate.FItemList(i).Fcouponbuyprice , mdchoicerotate.FItemList(i).Fbuycash)%></td>
				</tr>
				<tr>
					<td>��౸�� :</td>
					<td><%=fnColor(mdchoicerotate.FItemList(i).Fmwdiv,"mw")%>-<%=mdchoicerotate.FItemList(i).deliveryTypeName(mdchoicerotate.FItemList(i).Fdeliverytype)%></td>
				</tr>
				<tr>
					<td>�����Ȳ :</td>
					<td><a href="javascript:popCurrentStock('<%= mdchoicerotate.FItemList(i).Flinkitemid %>');">[����]</a></td>
				</tr>
			</table>
			<% end if %>
		</td>
		<td align="center">
			<input type="text" name="disporder" value="<%= mdchoicerotate.FItemList(i).FDisporder %>" size="3" style="text-align:center" class="text">
		</td>
		<td align="center" class="etcInfo">
			<%= formatdate(mdchoicerotate.FItemList(i).Fstartdate,"0000.00.00") %>
			<br>
			<%
				If cdate(mdchoicerotate.FItemList(i).Fstartdate) <= date() and  cdate(mdchoicerotate.FItemList(i).Fenddate) >= date()  Then
					Response.write " <span style=""color:red"">(������)</span>"					
				end If
			%>
		</td>
		<td align="center" class="etcInfo">
			<%= formatdate(mdchoicerotate.FItemList(i).Fenddate,"0000.00.00") %><br>
			<%
				If clng(datediff("d", now() , mdchoicerotate.FItemList(i).Fenddate)) < 0 Or clng(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(����)</span>"
				ElseIf cInt(datediff("d", mdchoicerotate.FItemList(i).Fenddate , now())) < 1  Then '���ó�¥
					If cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) < 24 Then ' ����
						Response.write " <span style=""color:red"">(�� "& cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
					End If 
				End If 
			%>
		</td>		
		<td align="center" class="etcInfo">
			<%= FormatDateTime(mdchoicerotate.FItemList(i).Fregdate,2) %>
		</td>
		<td align="center" class="etcInfo">
			<%
				If mdchoicerotate.FItemList(i).FLowestPrice = "Y" Then
					Response.write "���"
				Else
					Response.write "������"
				End If
			%>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).IsSoldOut then %>
			<font color="red"><%=mdchoicerotate.FItemList(i).FSellyn%></font>
			<% else %>
			<font color="blue"><%=mdchoicerotate.FItemList(i).FSellyn%></font>
			<% end if %>
		</td>
		<td align="center" class="etcInfo"><%= mdchoicerotate.FItemList(i).Fregname %></td>
		<td align="center" class="etcInfo"><%= mdchoicerotate.FItemList(i).Fworkername %></td>
		<td align="center" class="etcInfo"><button class="button" onclick="delItem('<%=mdchoicerotate.FItemList(i).Fidx %>');return false;">��ǰ����</button></td>
	</tr>
<% next %>
</tbody>
	<tr>
		<td colspan="13" align="center" bgcolor="white">
			<% if mdchoicerotate.HasPreScroll then %>
				<a href="javascript:NextPage('<%= mdchoicerotate.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + mdchoicerotate.StarScrollPage to mdchoicerotate.FScrollCount + mdchoicerotate.StarScrollPage - 1 %>
				<% if i>mdchoicerotate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if mdchoicerotate.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<%
set mdchoicerotate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->