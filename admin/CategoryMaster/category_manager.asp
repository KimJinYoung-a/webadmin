<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_contents_managecls.asp" -->

<%
dim research,isusing, fixtype, linktype, poscode, validdate, vCateCode, prevDate
dim page, cdl, cdm, imgSize
dim strParm

isusing = request("isusing")
research= request("research")
poscode = request("poscode")
fixtype = request("fixtype")
page    = request("page")
validdate= request("validdate")
prevDate = request("prevDate")
cdl		= request("cdl")
cdm		= request("cdm")
vCateCode = Request("catecode")

if ((research="") and (isusing="")) then
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1

strParm = "isusing="&isusing&"&poscode="&poscode&"&fixtype="&fixtype&"&validdate="&validdate&"&prevDate="&prevDate&"&catecode="&vCateCode
dim oposcode
set oposcode = new CCateContentsCode
oposcode.FRectPosCode = poscode

if (poscode<>"") then
    oposcode.GetOneContentsCode
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FPageSize = 10
oCateContents.FCurrPage = page
oCateContents.FRectIsusing = isusing
oCateContents.FRectfixtype = fixtype
oCateContents.FRectPosCode = poscode
oCateContents.FRectvaliddate = validdate
oCateContents.FRectSelDate = prevDate
oCateContents.FRectDisp1 = vCateCode
oCateContents.GetCateContentsList

dim i
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/categorymaster/popCatePosCodeEdit.asp','catePosCodeEdit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewCateContents(idx){
    var popwin = window.open('/admin/categorymaster/popCateContentsEdit.asp?idx=' + idx+'&<%=strParm%>','catePosCodeEdit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function AssignReal(){
    if (chkConfirm()) {
		 var popwin = window.open('','refreshFrm_Cate','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Cate";
		 <% If poscode <> "" Then %>
		<% If oposcode.FOneItem.Flinktype = "X" Then %>
			refreshFrm.action = "<%=wwwUrl%>/chtml/dispcate/main_make_xml.asp?poscode=<%=poscode%>&catecode=<%=vCateCode%>&term="+document.getElementById("vTerm").value+"";
		<% Else %>
			refreshFrm.action = "<%=wwwUrl%>/chtml/dispcate/make_cate_contents_JS.asp?poscode=<%=poscode%>&catecode=<%=vCateCode%>";
		<% End If %>
		<% End If %>
		 refreshFrm.submit();
	}
}

function AssignRealTest(){
    if (chkConfirm()) {
		 var popwin = window.open('','refreshFrm_Cate','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Cate";
		 <% If poscode <> "" Then %>
		<% If oposcode.FOneItem.Flinktype = "X" Then %>
			refreshFrm.action = "<%=wwwUrl%>/chtml_test/dispcate/main_make_xml.asp?poscode=<%=poscode%>&catecode=<%=vCateCode%>&term="+document.getElementById("vTerm").value+"";
		<% Else %>
			refreshFrm.action = "<%=wwwUrl%>/chtml_test/dispcate/make_cate_contents_JS.asp?poscode=<%=poscode%>&catecode=<%=vCateCode%>";
		<% End If %>
		<% End If %>
		 refreshFrm.submit();
	}
}

function AssignRealRightNow(idx){
    if (chkConfirm()) {
		 var popwin = window.open('','refreshFrm_Cate','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Cate";
		refreshFrm.action = "<%=wwwUrl%>/chtml/dispcate/catemain_linkbanner_make.asp?poscode=<%=poscode%>&catecode=<%=vCateCode%>&idx="+idx+"";
		 refreshFrm.submit();
	}
}

function AssignTest(cte){
		<% if application("Svr_Info")="Dev" then %>
		    var popwin = window.open("http://2015www.10x10.co.kr/shopping/category_main_test.asp?disp="+cte+"&chkTestDate="+document.getElementById("iSD").value,"_blank");
		<% else %>
		    var popwin = window.open("http://www1.10x10.co.kr/shopping/category_main_test.asp?disp="+cte+"&chkTestDate="+document.getElementById("iSD").value,"_blank");
		<% end if %>
	    popwin.focus();
}

function chkConfirm() {
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
		return false;
	}
	else{
		return true;
	}
}

function jsCateMainBrandItem(idx){
	var popupitem = window.open("/admin/categorymaster/category_brand_itempop.asp?idx="+idx+"", "popupitem", "width=1000,height=800,scrollbars=yes,resizable=yes");
	popupitem.focus();
}
</script>

<%
	If poscode = "370" Then
		Response.Write "<br><font color=red size=3><strong><u>�� 04 ī�װ� �귣���ȹ�� �Է��� ����ī�װ�, �귣����̵�, �귣����ī��, �ݿ��������� �� �ʼ��Է��Դϴ�. ��ǰ������ ������ �����ϴ�.</u></strong></font><br><br>"
	End IF
%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a"><input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������</td>
		<td class="a" >
			������ġ
			<% call DrawCatePosCodeCombo("poscode",poscode, "onChange='frm.submit();'") %>
			&nbsp;&nbsp;
			����ī�װ� :
			<%
			Dim cDisp
			SET cDisp = New cDispCate
			cDisp.FCurrPage = 1
			cDisp.FPageSize = 2000
			cDisp.FRectDepth = 1
'			cDisp.FRectUseYN = "Y"
			cDisp.GetDispCateList()

			If cDisp.FResultCount > 0 Then
				Response.Write "<select name=""catecode"" class=""select"" onChange=""frm.submit();"">" & vbCrLf
				Response.Write "<option value="""">����</option>" & vbCrLf
				For i=0 To cDisp.FResultCount-1
					Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
				Next
				Response.Write "</select>&nbsp;&nbsp;&nbsp;"
			End If
			Set cDisp = Nothing
			%>
			<br>
		    ��뱸��
			<select name="isusing" class="select" onChange="frm.submit();">
			<option value="">��ü
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
			<option value="N" <% if isusing="N" then response.write "selected" %> >������
			</select>
			&nbsp;&nbsp;
			���뱸��
			<% call DrawFixTypeCombo ("fixtype", fixtype, "") %>

			<% if poscode <> "" then %>
			<% If (oposcode.FOneItem.Flinktype = "X") Then %>
	        &nbsp;&nbsp;
	        �������� <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<% End If %>
			<% End If %>

			<% if C_ADMIN_AUTH then %>
			&nbsp;&nbsp;
			<input type="button" value="�ڵ����" onClick="popPosCodeManage();" class="button">
			<% end if %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
    <td colspan="4">
    <% if poscode <> "" and vCateCode <> "" then %>
    	<% If oposcode.FOneItem.Flinktype = "X" Then %>
		    ������ �����Ͽ� <input type="text" name="vTerm" id="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�
		    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> <b>Real ����(����)</b></a>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="javascript:AssignRealTest('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"><b> �׽�Ʈ ����</b></a>
			&nbsp;
			 <input id="iSD" name="iSD" value="<%=Left(now(), 10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ->
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "iSD", trigger    : "iSD_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
	//					CAL_End.args.min = date;
	//					CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<a href="" onclick="AssignTest('<%=vCateCode%>');return false;"><b>[�׽�Ʈ ������ Ȯ���ϱ�]</b></a>
		<% Else %>
			<!--<a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> <b>Real ����</b></a>-->
		<% End If %>
    <% end if %>
    </td>
    <td colspan="11" align="right"><a href="javascript:AddNewCateContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
    <td width="60">idx</td>
    <td width="80">ī�װ�</td>
    <td width="100">���и�</td>
    <td width="40">�̺�Ʈ<br>��ȣ</td>
    <td width="150">�̹���</td>
    <td width="50">��ũ<br>����</td>
    <td width="80">�ݿ�<br>�ֱ�</td>
    <td width="76">������</td>
    <td width="76">������</td>
    <td width="30">����<br>��ȣ</td>
    <td width="30">���<br>����</td>
    <td width="40">�����</td>
    <td width="40">�۾���</td>
    <td></td>
</tr>
<%
	for i=0 to oCateContents.FResultCount - 1

		'�̹��� ũ�� ����
		if oCateContents.FItemList(i).Fimagewidth>=oCateContents.FItemList(i).Fimageheight then
			if oCateContents.FItemList(i).Fimagewidth>=250 then
				imgSize = "width=250"
			else
				imgSize = ""
			end if
		else
			if oCateContents.FItemList(i).Fimageheight>=66 then
				imgSize = "height=66"
			else
				imgSize = ""
			end if
		end if
%>
<% if (oCateContents.FItemList(i).IsEndDateExpired) or (oCateContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= oCateContents.FItemList(i).Fidx %></td>
    <td align="center"><a href="?menupos=<%= request("menupos") %>&poscode=<%= poscode %>&catecode=<%=oCateContents.FItemList(i).Fdisp1%>"><%= oCateContents.FItemList(i).Fcodename %></a></td>
    <td align="center">
		<a href="?menupos=<%= request("menupos") %>&poscode=<%= oCateContents.FItemList(i).Fposcode %>&catecode=<%=vCateCode%>"><%= oCateContents.FItemList(i).Fposname %></a>
		<% If oCateContents.FItemList(i).Fevt_stdt <> "" Then %>
		<br/><br/>
		�̺�Ʈ �Ⱓ : <span style="color:red"><%=oCateContents.FItemList(i).Fevt_stdt %>~<%=oCateContents.FItemList(i).Fevt_etdt %></span>
		<% End If %>
	</td>
    <td align="center"><% if instr(oCateContents.FItemList(i).Flinkurl,"eventid")>0 then Response.Write Right(oCateContents.FItemList(i).Flinkurl,len(oCateContents.FItemList(i).Flinkurl)-Instr(oCateContents.FItemList(i).Flinkurl,"eventid=")-7) %></td>
    <td>
		<a href="javascript:AddNewCateContents('<%= oCateContents.FItemList(i).Fidx %>');">
			<% If Trim(oCateContents.FItemList(i).getImageUrl)="" Then %>
				<% If Trim(oCateContents.FItemList(i).FevtEtcImg)<>"" Then %>
					<img <%=imgSize%> src="<%= oCateContents.FItemList(i).FevtEtcImg %>" border="0">
				<% Else %>
					<img <%=imgSize%> src="<%= oCateContents.FItemList(i).FevtEtcBasicImg %>" border="0">
				<% End If %>
			<% Else %>
				<img <%=imgSize%> src="<%= oCateContents.FItemList(i).getImageUrl %>" border="0">
			<% End If %>
	    	<% If poscode = "370" Then %>[�������]<% End If %>
    	</a>
    	<% If poscode = "370" Then %>
    		&nbsp;&nbsp;<br />�귣��ID : <%=oCateContents.FItemList(i).Fmakerid%>
    		<a href="http://www.10x10.co.kr/street/street_brand_sub06.asp?makerid=<%=oCateContents.FItemList(i).Fmakerid%>" target="_blank">[����]</a>
    	<% End If %>
    </td>
    <td align="center"><%= oCateContents.FItemList(i).getlinktypeName %></td>
    <td align="center"><%= oCateContents.FItemList(i).getfixtypeName %></td>
    <td align="center"><%= oCateContents.FItemList(i).FStartdate %></td>
    <td align="center">
    <% if (oCateContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oCateContents.FItemList(i).FEnddate,10) %></font>
    <% else %>
    <%= Left(oCateContents.FItemList(i).FEnddate,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oCateContents.FItemList(i).FsortNo %></td>
    <td align="center"><%= oCateContents.FItemList(i).FIsusing %></td>
    <td align="center"><%= oCateContents.FItemList(i).Fregname %></td>
    <td align="center"><%= oCateContents.FItemList(i).Fworkername %></td>
    <td>
    <% if oCateContents.FItemList(i).Flinktype="L" then %>
    <a href="javascript:AssignRealRightNow('<%= oCateContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> <b>��� ����</b></a>
    <% else %>
    	<% If poscode = "365" OR poscode = "370" Then %>
    		<input type="button" value="��ǰ����" onClick="jsCateMainBrandItem('<%= oCateContents.FItemList(i).Fidx %>')">
    		<br><font color="red" size="3"><b>����:<%=oCateContents.FItemList(i).Fbrandcnt%>��</b></font>
    	<% End If %>
    <% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="15" align="center">
    <% if oCateContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCateContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oCateContents.StarScrollPage to oCateContents.FScrollCount + oCateContents.StarScrollPage - 1 %>
		<% if i>oCateContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oCateContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
<%
set oposcode = Nothing
set oCateContents = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
