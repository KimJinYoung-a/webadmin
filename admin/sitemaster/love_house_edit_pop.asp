<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/love_house_main_cls.asp" -->
<%

'// ������, 2015-07-07, skyer9
dbget.close()
response.end

dim objMon
Set objMon = Server.CreateObject("DEXT.FileUploadMonitor")
objMon.UseMonitor(True)
Set objMon = Nothing

dim page,ix,i,idx
page=request("page")
if page="" then page="1"
idx=request("idx")


dim loveitem
set loveitem = new LoveHouseOne
loveitem.FRectIdx=idx
loveitem.GetLoveMainOne
 %>
<script>

function subInput(){
	var frm;
	frm=document.inputfrm;

	strAppVersion = navigator.appVersion;
	if (frm.winimage.value != "") {
		if (strAppVersion.indexOf('MSIE')!=-1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4) {
			winstyle = "dialogWidth=385px; dialogHeight:150px; center:yes";
			window.showModelessDialog("show_progress.asp?nav=ie", null, winstyle);
		} else {
			winpos = "left=" + ((window.screen.width-380)/2) + ",top=((window.screen.height-110)/2)";
			winstyle = "width=380,height=110,status=no,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,copyhistory=no," + winpos;
			window.open("sshow_progress.asp",null,winstyle);
		}
	}

	frm.submit();
}
function showImg(v){

	v.imgpan.src=v.inputimg.value;
}
function showImg1(v){

	v.editpan.src=v.editimg.value;
}
function showImg2(v){
	v.winimgpan.src=v.winimage.value;
}
function subedit(frm){
	frm.mode.value="edit";
	frm.submit();
}

function subView(v) {
	document.tempfrm.idx.value=v;
	document.tempfrm.submit();
}

 </script>
<div align="left" valign="top">
<table width="650" cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
	<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/dolovehousewinner.asp" enctype="multipart/form-data">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="mode" value="edit">
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">��÷�� ���̵�</td>
		<td><input type="text" name="userid" size="20" value="<% =loveitem.Fuserid %>"></td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">��÷��</td>
		<td><input type="text" name="windate" size="20" value="<% =loveitem.Fwindate %>">(ex:2006-03-01)</td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">Title �̹���</td>
		<td>
			<input type="file" name="inputimg" size="35" Onchange="showImg(this.form)">
			<br>
			<img name="imgpan" src="<% =loveitem.FImage %>">
		</td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">��ũ</td>
		<td><input type="text" name="inputlink" size="35" value="<% = loveitem.Flink %>"></td>
	</tr>
	<tr class="a">
		<td align="center bgcolor="#DDDDFF""���� �̹���</td>
		<td>
			<input type="file" name="winimage" size="35" Onchange="showImg2(this.form)">
			<br>
			<img name="winimgpan" src="<% =loveitem.FWinImage %>"></td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">�̹��� ��</td>
		<td> �� �̸� �� <font color="blue">"lovemap"</font>���� �������ּ���<br>
			&lt;map name="lovemap"&gt;<br>
			&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('<font color="blue">��ǰ��ȣ</font>');" onfocus="this.blur();"&gt;<br>
			&lt;/map&gt; <br>
			<textarea name="lovemap" cols="80" rows="10"><% =loveitem.FLoveMap %></textarea></td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">WishList ��ǰ��ȣ</td>
		<td>
			��ǰ1: <input type="text" name="itemid1" size="10" maxlength="10" value="<% =loveitem.Fitemid1 %>"><br>
			��ǰ2: <input type="text" name="itemid2" size="10" maxlength="10" value="<% =loveitem.Fitemid2 %>"><br>
			��ǰ3: <input type="text" name="itemid3" size="10" maxlength="10" value="<% =loveitem.Fitemid3 %>"><br>
			��ǰ4: <input type="text" name="itemid4" size="10" maxlength="10" value="<% =loveitem.Fitemid4 %>"><br>
		</td>
	</tr>
	<tr class="a">
		<td align="center" bgcolor="#DDDDFF">��� ����</td>
		<td>
			<input type="radio" name="isusing" value="Y" <% if loveitem.FIsusing="Y" then response.write "checked" %>>���
			<input type="radio" name="isusing" value="N" <% if loveitem.FIsusing="N" then response.write "checked" %>>��� ����
		</td>
	</tr>
	<tr class="a">
		<td colspan="2" align="center"><input type="button" Onclick="javascript:subInput();" value="����"></td>
	</tr>
	</form>
</table>
<br><br><br>
</div>


<%
set loveitem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
