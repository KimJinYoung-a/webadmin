<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/openGift.asp
' Description :  ��ü�����̺�Ʈ ���� 369��.
' History : 2010.04 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<%
dim eCode : eCode=request("eC")
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, etag
Dim echkdisp, ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,ebimg,etemp,emimg,ehtml,eisort,eiaddtype,edid,emid,efwd,selPartner
Dim eusing, tmp_cdl, tmp_cdm, elktype, elkurl, ebimg2010, gimg, ebrand, eicon, ecommenttitle, elinkcode
dim dopendate, dclosedate, blnFull, blnIteminfo
dim ehtml5

dim eFolder : eFolder = eCode

Dim oOpenGift
Dim imod : imod="I"
Dim frontopen,  OGtitle, OopenHtml, OopenHtmlWeb, opengiftType, opengiftScope
set oOpenGift=new CopenGift
oOpenGift.FRectEventCode = eCode
oOpenGift.getOneOpenGift

if (oOpenGift.FResultCount>0) then
    frontopen = oOpenGift.FOneItem.FfrontOpen
    OGtitle   = oOpenGift.FOneItem.FopenImage1
    imod      = "E"
    eCode     = oOpenGift.FOneItem.Fevent_code
    OopenHtml = db2Html(oOpenGift.FOneItem.FopenHtml)
    OopenHtmlWeb = db2Html(oOpenGift.FOneItem.FopenHtmlWeb)
    opengiftType = oOpenGift.FOneItem.FopengiftType
    opengiftScope = oOpenGift.FOneItem.FopengiftScope
end if


IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ekind 		=	cEvtCont.FEKind
	eman 		=	cEvtCont.FEManager
	escope 		=	cEvtCont.FEScope
	selPartner	=	cEvtCont.FEPartnerID
	ename 		=	db2html(cEvtCont.FEName)
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	IF datediff("d",now,eeday) <0 THEN estate = 9 '�Ⱓ �ʰ��� ����ǥ��
	eregdate	=	cEvtCont.FERegdate
	eusing		= 	cEvtCont.FEUsing

	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	echkdisp 		= cEvtCont.FChkDisp
	tmp_cdl 		= cEvtCont.FECategory

	tmp_cdm		= cEvtCont.FECateMid
	esale 			= cEvtCont.FESale
	egift 			= cEvtCont.FEGift
	ecoupon 		= cEvtCont.FECoupon
	ecomment 		= cEvtCont.FECommnet
	ebbs 			= cEvtCont.FEBbs
	eitemps 		= cEvtCont.FEItemps
	eapply 			= cEvtCont.FEApply
	elktype			= cEvtCont.FELinkType
	IF elktype="" Then elktype="E" '//��ũŸ�� �⺻�� ����
	elkurl			= cEvtCont.FELinkURL
	ebimg 			= cEvtCont.FEBImg
	ebimg2010		= cEvtCont.FEBImg2010
	gimg			= cEvtCont.FEGImg
	etemp			= cEvtCont.FETemp
	if etemp = 5 or etemp = 6  THEN	'���۾� �̺�Ʈ �� ��� ó��
		ehtml5 		= db2html(cEvtCont.FEHtml)
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml)
	end if
	eisort 			= cEvtCont.FEISort
	edid 			= cEvtCont.FEDId
	emid 			= cEvtCont.FEMId
	efwd 			= db2html(cEvtCont.FEFwd)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode   	= cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
 	blnFull			= cEvtCont.FEFullYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)

	set cEvtCont = nothing
END IF

Dim arreventstate
arreventstate= fnSetCommonCodeArr("eventstate",False)


%>

<script language="javascript">
function saveOpenGift(){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frmOpenGift.submit();
    }
}

function jsLastEvent(){
  var winLast,eKind;
  eKind = 1;
  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >

<tr>
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ���� ���� �̺�Ʈ ���� </td>
</tr>

<tr>
	<td>
		<table width="1100" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ�ڵ�</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%= eCode %>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%=ename%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ�Ⱓ</B></td>
		   		<td bgcolor="#FFFFFF">
		   			������ : <%=esday%>
		   			~ ������ : <%=eeday%>
		   		</td>
		   	</tr>

		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ����</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <%= Replace(fnGetCommCodeArrDesc(arreventstate,estate),"���¿���","����") %>

		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  ����ó���� : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  ����ó���� : <%=dclosedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		</table>
	</td>
</tr>
<tr>
    <td>
    <table width="1100" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmOpenGift" method="post"  action="openGift_process.asp" onSubmit="return jsEvtSubmit(this);">
    <input type="hidden" name="imod" value="<%= imod %>">
    <input type="hidden" name="menupos" value="<%=menupos%>">
    <input type="hidden" name="OGtitle" value="<%=OGtitle%>">
    <input type="hidden" name="eCode" value="<%=eCode%>">

		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>���¿���</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <input type="radio" name="frontopen" value="Y" disabled >Open
		   			<input type="radio" name="frontopen" value="N" checked >Close
		   			(�űԵ�Ͻÿ��� Close�� �����˴ϴ�.)
		   		    <% else %>
		   		    <input type="radio" name="frontopen" value="Y" <%= chkIIF(frontopen="Y","checked","") %> >Open
		   			<input type="radio" name="frontopen" value="N" <%= chkIIF(frontopen="N","checked","") %> >Close
		   			<% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   	    <td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>��ü ���� ����</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <input type="radio" name="opengiftType" value="1" checked >��ü���� �̺�Ʈ
		   		    <input type="radio" name="opengiftType" value="9"  >���̾ �̺�Ʈ
		   		    <% else %>
		   		    <input type="radio" name="opengiftType" value="1" <%= chkIIF(opengiftType=1,"checked","") %> >��ü���� �̺�Ʈ
		   		    <input type="radio" name="opengiftType" value="9" <%= chkIIF(opengiftType=9,"checked","") %> >���̾ �̺�Ʈ
		   		    <% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   	    <td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>���� ����</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <label><input type="radio" name="opengiftScope" value="1" checked >��ü</label>
		   		    <label><input type="radio" name="opengiftScope" value="3"  >�����</label>
		   		    <label><input type="radio" name="opengiftScope" value="5"  >APP</label>
		   		    <% else %>
		   		    <label><input type="radio" name="opengiftScope" value="1" <%= chkIIF(opengiftScope="1","checked","") %> >��ü</label>
		   		    <label><input type="radio" name="opengiftScope" value="3" <%= chkIIF(opengiftScope="3","checked","") %> >�����</label>
		   		    <label><input type="radio" name="opengiftScope" value="5" <%= chkIIF(opengiftScope="5","checked","") %> >APP</label>
		   		    <% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>Ÿ��Ʋ�̹���</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <input type="button" class="button" value="�̹������" onClick="jsSetImg('<%=eFolder%>','<%=OGtitle%>','OGtitle','spantitle')">
		   		    (��ٱ��Ͽ� ǥ�õǴ� �̹���)
		   		    <div id="spantitle" style="padding: 5 5 5 5">
		   				<%IF OGtitle <> "" THEN %>
		   				<img  src="<%=OGtitle%>" width="100%" />
		   				<a href="javascript:jsDelImg('OGtitle','spantitle');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>

		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>���� ��������<br>��<br>���ǻ���</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<textarea name="openHtmlWeb" cols="120" rows="10"><%=OopenHtmlWeb%></textarea>
<p><font color="blue">�����ڵ�</font></p>
<textarea name="openHtmlWeb_Sample" cols="120" rows="8" style="border:0">
<li>3/7/15���� �̻� ���Ž� ���ϸ���, ����, ����ī�� ���� ��� �� ���� �����ݾ� ����</li>
<li class="tMar07">�ٹ����� ��ۻ�ǰ�� �����Ͽ� 3/7/15�������� ���Ž� ��ǰ �Ǵ� ���� �߿� ����</li>
<li class="tMar07">�ٹ����� ��ۻ�ǰ ���� ��ü��ǰ������ 7/15�������� ���Ž� ����ǰ�� ������ ���ð����մϴ�.<br />(3�����̻� ���Ž� ������ ��������ǰ�� �����ϴ�.)</li>
<li class="tMar07">����ǰ ������ 10�� 26�� �ϰ� �߱� �ص帳�ϴ�.</li>
<li class="tMar07">����ǰ ��ǰ�� �ٸ� ������� �޴°��� �Ұ��մϴ�.</li>
<li class="tMar07">�÷��� �������� �߼۵Ǹ� ��ȯ�� �Ұ��մϴ�.</li>
</textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>����Ͽ� ��������<br>��<br>���ǻ���</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<textarea name="openHtml" cols="140" rows="25"><%=OopenHtml%></textarea>
<p><font color="blue">�����ڵ�</font></p>
<textarea name="openHtml_Sample" cols="140" rows="25"  style="border:0">
<div id="lyGiftNoti" style="display:none;">
	<div class="layerPopup lyGiftNoti">
		<dl>
			<dt>���ǻ���</dt>
			<dd>
				<ul class="cartInfoV16a">
					<li>���ϸ���, ���� ���� ��� �� ����Ȯ���ݾ� ����</li>
					<li>�ٹ����� ��ۻ�ǰ���� 2�����̻� ���Ž� ����ǰ ����</li>
					<li>����ǰ�� �ٹ����� ��ۻ�ǰ�� �Բ� ���</li>
					<li>ȯ�� �� ��ȯ���� ���� �������� �̴� ��, ����ǰ �� ���ϸ��� ��ǰ �ʼ�</li>
					<li>����ǰ�� �������� �߼۵Ǹ� ��ȯ �Ұ�</li>
					<li>����ǰ ������ �̺�Ʈ ����</li>
				</ul>
			</dd>
		</dl>
		<button class="lyClose" onclick="fnClosePartLayer();">�ݱ�</button>
	</div>
</div>
<div class="bxLGy2V16a grpTitV16a">
	<h2>[������ ����� ���� ���� ���� �ְ� ����] ����ǰ ����</h2>
	<i class="icoQuestV16a" onClick="fnOpenPartLayer();return false;">����ǰ ���� ���ǻ���</i>
</div>
<div class="bxWt1V16a freebieSltV16a">
	<div class="bxWt1V16a">2016.03.12~ 2016.03.28 (���� ������ ����)</div></textarea>
		   		</td>
		   	</tr>
    </table>
    </form>
    </td>
</tr>
<tr>
	<td width="800" height="40" align="right">
		<img src="/images/icon_save.gif" onClick="saveOpenGift()"  style="cursor:pointer">
		<a href="/admin/eventmanage/event/openGift.asp?menupos=1184"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>

</table>
<%
set oOpenGift=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->