<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  [2016 S/S ����] Wedding Membership �������� ���ε� ����Ʈ
' History : 2016.09.12 ���¿�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/datamart/mkt/73007_manageCls.asp"-->
<%
	Dim o73007, i , page , state ,idx, evt_code, imgurl
'	dim userid, suserid, isusing
	menupos = request("menupos")
	page = request("page")
	state = request("state")

	if page = "" then page = 1

	IF application("Svr_Info") = "Dev" THEN
		evt_code   =  66201
	Else
		evt_code   =  73007
	End If

If session("ssBctId") ="greenteenz" Or session("ssBctId") = "djjung" Then
else
	response.write "�����ڸ� �� �� �ִ� ������ �Դϴ�"
	response.End
end if

imgurl = staticImgUrl&"/contents/contest/"&evt_code&"/"

set o73007 = new CMagaZineContents
	o73007.FPageSize = 20
	o73007.FCurrPage = page
	o73007.FRectevt_code = evt_code
	o73007.fnGet73007subevtList()
%>
<script type="text/javascript">
function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

//�̹��� Ȯ�� ��â���� �����ֱ�
function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}

//���� Y,N �˻�
function chggubun(val){
	parent.location.href="/admin/datamart/mkt/73007_manage.asp?state="+val;
}

function reggubunok(evtcode,gidx,uid){
	gubunokfrm.action="/admin/datamart/mkt/73007proc.asp";
	gubunokfrm.mode.value="gubunok";
	gubunokfrm.evt_code.value=evtcode;
	gubunokfrm.sub_idx.value=gidx;
	gubunokfrm.userid.value=uid;
	gubunokfrm.target="evtFrmProc";
	gubunokfrm.submit();
	return;
}

</script>

<form name="frm" method="post" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>

<%
If session("ssBctId") ="greenteenz" Or session("ssBctId") = "djjung" Then
%>
	<!-- ����Ʈ ���� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
	<tr bgcolor="#FFFFFF">
		<td colspan="20">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					�˻���� : <b><%= o73007.FTotalCount %></b>
					&nbsp;
					������ : <b><%= page %> / <%=  o73007.FTotalpage %></b>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="5%">idx</td>
		<td width="5%">�̺�Ʈ�ڵ�</td>
		<td width="5%">userid</td>
		<td width="15%">���̹���</td>
	</tr>
	<% if o73007.FresultCount > 0 then %>
		<% for i=0 to o73007.FresultCount-1 %>
			<%
'			if isarray(split(o73007.FItemList(i).Fsub_opt1,"/!/")) then
'				userid = split(o73007.FItemList(i).Fsub_opt1,"/!/")(0)
'				suserid = split(o73007.FItemList(i).Fsub_opt1,"/!/")(1)
'				isusing = split(o73007.FItemList(i).Fsub_opt1,"/!/")(2)
'			end if
			%>
		<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
			<td align="center"><%= o73007.FItemList(i).Fidx %></td>
			<td align="center"><%= o73007.FItemList(i).Fevt_code %></td>
			<td align="center"><%= o73007.FItemList(i).Fuserid %></td>
			<td align="center"><img src="<%= imgurl %><%= o73007.FItemList(i).Fsub_opt3 %>" width=70 border=0 onclick="showimage('<%= imgurl %><%= o73007.FItemList(i).Fsub_opt3 %>');" style="cursor:pointer;"></td>
		</tr>
		<% Next %>
		<form name="gubunokfrm" method="post" action="" style="margin:0px;">
		<input type="hidden" name="mode">
		<input type="hidden" name="evt_code">
		<input type="hidden" name="sub_idx">
		<input type="hidden" name="userid">
		</form>
	<tr>
		<td colspan="20" align="center" bgcolor="#FFFFFF">
		 	<% if o73007.HasPreScroll then %>
				<a href="javascript:NextPage('<%= o73007.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for i=0 + o73007.StartScrollPage to o73007.FScrollCount + o73007.StartScrollPage - 1 %>
				<% if i>o73007.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
			<% if o73007.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% end if %>
	</table>
<%
Else
	response.write "�����ڸ� �� �� �ִ� ������ �Դϴ�"
	response.End
End If
%>
<iframe id="evtFrmProc" name="evtFrmProc" src="about:blank" frameborder="0" width=0 height=0></iframe>
<% set o73007 = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->