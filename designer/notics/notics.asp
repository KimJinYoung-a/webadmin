<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim isNewContractTypeExists, noConfirmedCtrExists

noConfirmedCtrExists = isNotFinishNewContractExists(session("ssBctID"), session("ssGroupid"), isNewContractTypeExists)

%>

<script language="JavaScript">
<!--
function PopNotice(v){
    var popwin = window.open("/designer/notics/notics_read.asp?idx=" + v ,"PopNotice","width=1100,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function NextPage(ipage){
    document.searchform.target="_notice";
    document.searchform.action="/designer/notics/popnoticslist.asp";
	document.searchform.page.value= ipage;
	document.searchform.submit();

	document.searchform.target="";
	document.searchform.action="";
}

// ��Ű ����
var cookiedata = document.cookie;
function setCookie( name, value, expiredays) {
	var todayDate = new Date();
	var dom = document.domain;
	var _domain = "";
	if(dom.indexOf("10x10.co.kr") > 0){
		_domain = "10x10.co.kr";
	}
	todayDate.setDate( todayDate.getDate() + expiredays );
	document.cookie = name + "=" + escape( value ) + "; domain="+_domain+"; path=/; expires=" + todayDate.toGMTString() + ";"
}

// ������ �ε��� ��Ʈ�� ǥ��
window.onload = function()
{
<% if datediff("d","2013-10-31", date())<=0 then %>
	if ( cookiedata.indexOf("chkIntroTGV=done") < 0 ){
		switchIntro();
	}
<% end if %>
<% if datediff("d","2013-08-31", date())<=0 then %>
	if ( cookiedata.indexOf("chk2013TGV=done") < 0 ){
		switchIntro2();
	}
<% end if %>

<% if datediff("d","2016-10-31", date())<=0 then %>
	if ( cookiedata.indexOf("chk2016TGV=done") < 0 ){
		switchIntro3();
	}
<% end if %>

<% if (noConfirmedCtrExists) then %>
    //if ( cookiedata.indexOf("chk2013CTR=done") < 0 ){
		switchIntroCtr();
	//}
<% end if %>
}
// ��Ʈ�� On/Off
function switchIntro() {
	if(document.getElementById("2009TGV").style.display=='none') {
		document.getElementById("2009TGV").style.display=''
	} else {
		document.getElementById("2009TGV").style.display='none'
		setCookie( "chkIntroTGV", "done" , 1 );
	}
}

function switchIntro2() {
	if(document.getElementById("2013TGV").style.display=='none') {
		document.getElementById("2013TGV").style.display=''
	} else {
		document.getElementById("2013TGV").style.display='none'
		setCookie( "chk2013TGV", "done" , 1 );
	}
}

function switchIntro3() {
	if(document.getElementById("2016TGV").style.display=='none') {
		document.getElementById("2016TGV").style.display=''
	} else {
		document.getElementById("2016TGV").style.display='none'
		setCookie( "chk2016TGV", "done" , 1 );
	}
}

function switchIntroCtr() {
	if(document.getElementById("2016CTR").style.display=='none') {
		document.getElementById("2016CTR").style.display=''
	} else {
		document.getElementById("2016CTR").style.display='none'
		//setCookie( "chk2013CTR", "done" , 1 );
	}
}
//-->
</script>
<%

response.expires = 0

dim ibalju
dim yyyy1,mm1,dd1,nowdate
dim mibaljuCount, mibeasongCount

nowdate = Left(CStr(now()),10)
yyyy1 = Left(nowdate,4)
mm1   = Mid(nowdate,6,2)
dd1   = Mid(nowdate,9,2)


set ibalju = new CBaljuMaster

ibalju.FRectRegStart = DateSerial(yyyy1,mm1-1, dd1)
ibalju.FRectRegEnd = DateSerial(yyyy1,mm1, dd1+1)
ibalju.FRectDesignerID = session("ssBctID")
'ibalju.DesignerDateMiBaljuCount

if ibalju.FResultCount>0 then
	mibaljuCount = ibalju.FMasterItemList(0).FTotalea
else
	mibaljuCount = 0
end if

'ibalju.DesignerDateMiBeasongCount
if ibalju.FResultCount>0 then
	mibeasongCount = ibalju.FMasterItemList(0).FTotalea
else
	mibeasongCount =0
end if

	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent, strtitle
	Dim nInstr,searchmode,search,searchString
    Dim nboard
	Dim nboardFix

	pgsize = requestCheckVar(Request("pgsize"),10)
	if pgsize="" then
		pgsize = 10
	end if

	page = requestCheckVar(Request("page"),10)
	if page = "" then
		page = 1
	end if

set nboardFix = new CBoard
nboardFix.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboardFix.FRectFixonly = "on"
nboardFix.FPageSize = 7
nboardFix.FRectDesignerID = session("ssBctID")
nboardFix.design_notice_dispcate

set nboard = new CBoard
nboard.FRectFixonly = "off"
 
if Request("SearchMode") = "search" then
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
'nboard.FRectDesignerID = designer
nboard.FPageSize = pgsize
nboard.FCurrPage = page
'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
nboard.FRectsearch = request("search")
nboard.FRectsearch2 = request("SearchString")
nboard.FRectDesignerID = session("ssBctID")
nboard.design_notice_dispcate
else
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
'nboard.FRectDesignerID = designer
nboard.FPageSize = pgsize
nboard.FCurrPage = page
'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
'nboard.FRectOrderSerial = orderserial
nboard.FCurrPage = page
nboard.FRectDesignerID = session("ssBctID")
nboard.design_notice_dispcate
end if


dim sqlstr, csnofincnt, itemqanotfinish
csnofincnt = 0
itemqanotfinish = 0

sqlstr = "select count(id) as cnt"
sqlstr = sqlstr + " from [db_cs].[dbo].tbl_as_list c"
sqlstr = sqlstr + " where deleteyn='N'"
sqlstr = sqlstr + " and divcd not in ('5','7')"
sqlstr = sqlstr + " and c.currstate='1'"
sqlstr = sqlstr + " and makerid='" + session("ssBctID") + "'"

'rsget.Open sqlStr,dbget,1
'	csnofincnt = rsget("cnt")
'rsget.Close

'sqlstr = "select count(m.id) as cnt from [db_cs].[dbo].tbl_myqna m"
'sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item i on m.itemid=i.itemid"
'sqlstr = sqlstr + " where m.isusing = 'Y'"
'sqlstr = sqlstr + " and m.replyuser = ''"
'sqlstr = sqlstr + " and m.itemid>0"
'sqlstr = sqlstr + " and i.makerid='" + session("ssBctID") + "'"
'rsget.Open sqlStr,dbget,1
'	itemqanotfinish = rsget("cnt")
'rsget.Close

%>
<!-- // ��Ʈ�� ���̾� ���� // -->
<div id="2009TGV" style="position:absolute; width:500px; margin-top:-20px; margin-left:80px; display:none">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td style="padding:3px" bgcolor="#505050">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<!-- �귣�� ������ ��� �ȳ� -->
			<td height="480"><img src="<%= fixImgUrl %>/web2013/cscenter/notice/brand_admin_notice_20131010.jpg" border="0" width="500" height="480" usemap="#tgvmap" />
				<map name="tgvmap">
					<area shape="rect" coords="457,21,486,51" href="" onclick="document.getElementById('2009TGV').style.display='none'; return false;" alt="�ݱ�" />
					<area shape="rect" coords="165,392,334,441" href="/designer/notics/notics_read.asp?idx=626" target="_blank" alt="��������" />
				</map>
			</td>
			<!-- ��ǰ������ <td height="800"><img src="<%= fixImgUrl %>/web2011/cscenter/cs_message_info.gif" width="500" height="800"></td>-->
			<!-- �ޱⰡ�� ��ǰ������ <td height="288"><img src="<%= fixImgUrl %>/web2011/cscenter/cs_message_holiday.gif" width="500" height="800"></td>-->
			<!-- �߼� �λ� <td height="436"><img src="<%= fixImgUrl %>/web2009/main/pop_2009_tgv_designer.jpg" width="427" height="436"></td>-->
			<!-- ����ں��� ���� <td height="288"><img src="<%= fixImgUrl %>/pop_notice.gif" width="341" height="288"></td>-->
			<!-- 2012�� ������ ���� <td height="288"><img src="<%= fixImgUrl %>/web2011/cscenter/cs_message.gif" width="500" height="650"></td>-->
			<!-- ������� ���� <td height="288"><img src="<%= fixImgUrl %>/web2011/cscenter/cs_message_dlv.gif" width="500" height="650"></td>-->
		</tr>
		<tr>
			<td bgcolor="#C0C0C0" align="right" style="padding:2px;">�����Ϸ� ���� �ʱ� <input type="checkbox" onClick="switchIntro();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>
<div id="2013TGV" style="position:absolute; width:680px; margin-top:-20px; margin-left:140px; display:none;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td style="padding:3px" bgcolor="#505050">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<!-- ����ں��� ���� -->
			<td height="288"><img src="<%= fixImgUrl %>/web2011/cscenter/cs_message_categorymd.gif" width="680" height="1320"></td>
		</tr>
		<tr>
			<td bgcolor="#C0C0C0" align="right" style="padding:2px;">�����Ϸ� ���� �ʱ� <input type="checkbox" onClick="switchIntro2();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>

<div id="2016TGV" style="position:absolute; width:450px; margin-top:-20px; margin-left:80px; display:none;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td style="padding:3px" bgcolor="#505050">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<!-- ��Ʈ�ʻ� ������ ��ȣ �ȳ� ���� -->
			<td>
				<div style="padding:12px 5px; text-align:center; background-color:#e8e8e8; font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; border:1px dashed #ddd">
					<strong style="font-size:16px;">��Ʈ�ʻ� ������ ��ȣ �ȳ�</strong></span>
					<div style="background-color:#fff; padding:10px; margin-top:10px;">
						<div style="padding:10px;">
							<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:20px; color:#00cccc;">��Ʈ�ʻ� ������ :</strong>
							<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:26px; color:#00cccc; text-shadow:1px 1px rgba(0,51,51,0.4);">070-4868-1799</strong>
						</div>
						<div style="padding:5px;">
							<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:20px; color:#000;">���ֹ��� ���õ� ���Ǵ� �� ��ȣ�� �����ֽñ� �ٶ��ϴ�.</strong>
						</div>
						<div style="padding:10px;">
							<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:16px; color:#555;">(�̺�Ʈ, ������ �� ���� ���ù��Ǵ� ��� ���𿡰� �������ּ���.)</strong>
						</div>
					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#C0C0C0" align="right" style="padding:2px;">�����Ϸ� ���� �ʱ� <input type="checkbox" onClick="switchIntro3();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>

<div id="2016CTR" style="position:absolute; width:680px; margin-top:-20px; margin-left:140px; display:none;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td style="padding:3px" bgcolor="#505050">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a" bgcolor="#FFFFFF">
		<tr>
			<!-- ��༭ OPEN ���� -->
			<td align="center">
			<table cellspacing="0" cellpadding="0" style="border:0; width:760px; padding:0;">
            <tbody>
            <tr>
            	<td><img width="600" height="60" src="<%= fixImgUrl %>/web2008/mail/mail_header.gif" /></td>
            </tr>
            <tr>
            	<td style="border:5px solid #eee; padding:10px; background:#fff;">
            		<table cellspacing="0" cellpadding="0" style="width:100%; padding:0; margin:0" >
            		<tbody>
            		<tr>
            			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; line-height:1.6; padding:0; margin:0"><strong>�ȳ��ϼ���. �ٹ����� �Դϴ�.</strong><br />
            			    �ű� ��༭�� �߼� �Ǿ����ϴ�.<br /> 
            				2016�� 3������ �ٹ����� �ŷ��⺻��༭�� ����Ǿ����ϴ�.<br />
            				�ֱ� ��༭ �ŷ��Ͻ� ��� ��ü�� �ش�Ǵ� �κ��̴� ���ŷο�ô��� ȸ�� �� ��Ź�帳�ϴ�.<br /> 
            				
            				���ο� ��༭ ������ Ȯ�� �Ͻ� ��, ���� �ϼż� ����ڿ��� �������� �߼��� �ֽñ�ٶ��ϴ�.<br />
            				�� ��, ������ ���� �����ŵ� �Ǹ� �����ֽô� ��༭�� �ٹ����� ���������� 1�θ� �����ֽø� �˴ϴ�.<br />
            			</td>
            		</tr>
            		<tr>
            			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
            				<strong>* ��������</strong><br />
            				&nbsp;&nbsp;&nbsp;1.��༭ �ٿ�ε� / �� 2�� ���<br />
            				&nbsp;&nbsp;&nbsp;2.���޻翡�� ��༭ Ȯ���� ���� (���� ���ʿ�)/ 1�� ����߼�<br />
            				&nbsp;&nbsp;&nbsp;(pdf ������ ���÷��� pdf �� �ʿ��մϴ�. �� ���°�� ���� ��ũ���� �ٿ�ε� �����մϴ�. <a target="_blank" href="http://software.naver.com/software/summary.nhn?softwareId=MFS_100032" style="color:#333;">��pdf reader</a> , <a target="_blank" href="http://get.adobe.com/kr/reader/" style="color:#333;">adobe reader</a>)
            			</td>
            		</tr>
            		<tr>
            			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
            				<strong>* �����ֽ� ����</strong><br />
            				[���� ��ü��� / Ư����� ����ü]<br />
            				
            				&nbsp;&nbsp;&nbsp;- �ŷ��⺻��༭ 1��(ȸ��������� �ۼ� �ϹǷ� �� ȸ��� 1�� ����)<br />
            				&nbsp;&nbsp;&nbsp;- �ŷ��⺻���μ����Ǽ� ��1��(�귣�� ���̵𺰷� �ۼ��ϹǷ� �귣�� ���̵�� 1�� ����)<br />
            				&nbsp;&nbsp;&nbsp;- ���޻� ���� �������� ���� 1�� (�ŷ��⺻��༭ �ǵ��� ����)<br />
            				&nbsp;&nbsp;&nbsp;- �������� �纻 1��<br />
            				&nbsp;&nbsp;&nbsp;- ����� ����� �纻 1��<br /> 
            				 
            				[���� ���԰���ü]<br />
            				&nbsp;&nbsp;&nbsp;- �����԰�༭<br />
                            &nbsp;&nbsp;&nbsp;- �������� �纻 1��<br />
            				&nbsp;&nbsp;&nbsp;- ����� ����� �纻 1��<br /> 
            			</td>
            		</tr>
            		<tr>
            			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
            				<strong>* ��༭ �߰� �� �������</strong><br />  
            				&nbsp;&nbsp;&nbsp;- ��� ���� ��Ī��� ����: <strong>���� ���޻翡�� ���»�� ����</strong><br />
            				&nbsp;&nbsp;&nbsp;-  �г�Ƽ ����: <strong>�ŷ��⺻��� �μ����Ǽ� ��2��(��ϻ�ǰ�����) ����</strong><br />
            				&nbsp;&nbsp;&nbsp;- ��ǰ �Ǹ� �ߴ�: <strong>�ŷ��⺻��� �μ����Ǽ� ��2��(��ϻ�ǰ�����) ����</strong><br />
            				&nbsp;&nbsp;&nbsp;- ���Ⱓ: <strong>����Ϸ� ���� Ư�����ڱ���(�������� ����) 3���� ������ �ڵ� ����</strong><br />
            				&nbsp;&nbsp;&nbsp;-��� ����/����: <strong>�ŷ��⺻��༭ �� 19�� ����</strong><br />
            			</td>
            		</tr>
            	    <tr>
            			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
            				<strong>* ��༭ �����ֽǰ�</strong><br />
            				&nbsp;&nbsp;&nbsp;- �ּ� : (03086) ����� ���α� ���з�12�� 31 �������� 6�� �ٹ����� ���»� ��༭ ����� ��
            			</td>
            		</tr>
            		<tr>
            			<td style="padding:10px 0; margin:0;" align="right">
            				<a href="/designer/company/contract/ctrListBrand.asp?menupos=1623"><font color="blue">��༭ �ٿ�ε�&gt;&gt;</font></a>
            			</td>
            		</tr>
            		</table>
            	</td>
            </tr>
            </tbody>
            </table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#C0C0C0" align="right" style="padding:2px;">�ݱ� <input type="checkbox" onClick="switchIntroCtr();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>
<!-- // ��Ʈ�� ���̾� �� // -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<iframe onload="this.style.height=this.contentWindow.document.body.scrollHeight;" src="iiframesumary.asp" width="100%" height="105" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
		</td>
	</tr>
</table>

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="searchform"  method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="SearchMode" value="search">

	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="4">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ü��������</b>
	    </td>
	</tr>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<select class="select" name="search" size="1">
		   		<option value="title">������</option>
		   		<option value="name">�̸�</option>
		   		<option value="content">����</option>
			</select>
			<input name="SearchString" class="text" type="text">
			<input type="image" src="/images/icon_search.gif" width="45" height="20" border="0" align="absbottom"></a>
			�˻���� : <b><% = nboard.FTotalCount %></b>
		</td>
	</tr>
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">��ȣ</td>
    	<td>����</td>
      	<td width="100">�ۼ���</td>
      	<td width="100">�ۼ���</td>
    </tr>

	<form name="qnaform" method="post">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<% for ix=0 to nboardFix.FResultCount -1 %>
	<tr class="a" bgcolor="<%= adminColor("pink") %>">
		<td align="center" height="16">[����]</td>
		<td align="center"><a href="javascript:PopNotice('<%= nboardFix.BoardItem(ix).FRectIdx  %>');"><%= nboardFix.BoardItem(ix).FRectTitle %></a>
		<% if datediff("d",nboardFix.BoardItem(ix).Fregdate,now())<8 then %>
		&nbsp;<font color=red><b>new</b></font>
		<% end if %>
		</td>
		<td align="center"><%= nboardFix.BoardItem(ix).FRectName %></td>
		<td align="center"><%= FormatDateTime(nboardFix.BoardItem(ix).Fregdate,2) %></td>
	</tr>
	<% next %>
	<% if (nboard.FResultCount < 1) and (nboardFix.FResultCount < 1) then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center">[�������׿� ���� �����ϴ�.]</td>
	</tr>
	<% else %>
	<% for ix=0 to nboard.FResultCount -1 %>
	<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="22"><%= nboard.BoardItem(ix).FRectIdx  %></a></td>
		<td align="center">
		    <a href="javascript:PopNotice(<%= nboard.BoardItem(ix).FRectIdx %>);"><%= nboard.BoardItem(ix).FRectTitle %></a>
		    <% if datediff("d",nboard.BoardItem(ix).Fregdate,now())<8 then %>
		    &nbsp;<font color=red><b>new</b></font>
		    <% end if %>
		</td>
		<td align="center"><%= nboard.BoardItem(ix).FRectName %></td>
		<td align="center"><%= FormatDateTime(nboard.BoardItem(ix).Fregdate,2) %></td>
	</tr>
	<% next %>
	<% end if %>
	</form>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if nboard.HasPreScroll then %>
				<a href="javascript:NextPage('<%= nboard.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + nboard.StartScrollPage to nboard.FScrollCount + nboard.StartScrollPage - 1 %>
				<% if ix>nboard.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if nboard.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>



<%
set ibalju = Nothing
set nboardFix = Nothing
set nboard = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
