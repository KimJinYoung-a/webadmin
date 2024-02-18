<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	: 2013.09.30 ������ ����
'			  2022.07.04 �ѿ�� ����(isms���������)
'	Description : �ſ�ī�� ���θ�� ����(������ ������ display)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pgPromotionCls.asp"-->
<%
Dim idx : idx = requestCheckVar(getNumeric(request("idx")),10)
Dim sDt, eDt
Dim pgprogbn, cardcd, cimage, isusing, conts, regdate, contlink

Dim oCardPromo
SET oCardPromo= new CCardPromotion
oCardPromo.FRectIdx=idx
if (idx<>"") then
oCardPromo.getCardPromotionOne
end if

if oCardPromo.FResultCount>0 then
    cimage = oCardPromo.FOneItem.Fcimage
    pgprogbn = oCardPromo.FOneItem.Fpgprogbn
    cardcd = oCardPromo.FOneItem.FCardCd
    sDt = Left(oCardPromo.FOneItem.FSDt,10)
    eDt = Left(oCardPromo.FOneItem.FEDt,10)
    conts = ReplaceBracket(oCardPromo.FOneItem.Fconts)
    contlink = ReplaceBracket(oCardPromo.FOneItem.Fcontlink)
    isusing = oCardPromo.FOneItem.FIsUsing
    regdate = oCardPromo.FOneItem.FRegDate
end if
SET oCardPromo= Nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script type='text/javascript'>

function jsConfirmSm(){
    var frm = document.frmReg;

    if (frm.pgprogbn.value.length<1){
        alert('������ ���� �ϼ���.');
        frm.pgprogbn.focus();
        return false;
    }

    if ((frm.pgprogbn.value=='m')&&(frm.cardcd.value.length<1)){
        alert('ī��縦 ���� �ϼ���.');
        frm.cardcd.focus();
        return false;
    }

    if (frm.sDt.value.length<1){
        alert('�������� �Է� �ϼ���.');
        frm.sDt.focus();
        return false;
    }

    if (frm.eDt.value.length<1){
        alert('�������� �Է� �ϼ���.');
        frm.eDt.focus();
        return false;
    }

    if ((frm.pgprogbn.value!='m')&&(frm.cimage.value.length<1)){
        alert('�̹����� ���� �ϼ���.');
        return false;
    }

    if ((frm.pgprogbn.value=='m')&&(frm.conts.value.length<1)){
        alert('������ �Է� �ϼ���.');
        frm.conts.focus();
        return false;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        return true;
    }
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function chgpgprogbn(comp){
    var frm=comp.form;

    if (comp.value=="m"){
        $('#iimgtr').hide();
    }else{
        $('#iimgtr').show();
    }
}
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->
		<form name="frmReg" method="post" action="pop_RegPgPromotion_process.asp" onSubmit="return jsConfirmSm();" >
		<input type="hidden" name="idx" value="<%=idx%>">
		<input type="hidden" name="cimage" value="<%=cimage%>">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>
			<td><b>ī�� ���θ�� ��� �� ����</b></td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<% IF idx <> "" THEN%>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ��ȣ</td>
					<td bgcolor="#FFFFFF"><%=idx%></td>
				</tr>
				<% end if %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">����</td>
					<td bgcolor="#FFFFFF">
                    <% Call DrawSelectBoxCardPromoGubun("pgprogbn",pgprogbn,"onChange='chgpgprogbn(this)'") %>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">ī���</td>
					<td bgcolor="#FFFFFF">
                    <% Call DrawSelectBoxCardList("cardcd",cardcd) %>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�Ⱓ</td>
					<td bgcolor="#FFFFFF">
                    <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
                    <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
                    <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
                   <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
                    <script language="javascript">
                        var CAL_Start = new Calendar({
            				inputField : "sDt", trigger    : "sDt_trigger",
            				onSelect: function() {
            					var date = Calendar.intToDate(this.selection.get());
            					CAL_End.args.min = date;
            					CAL_End.redraw();
            					this.hide();
            				}, bottomBar: true, dateFormat: "%Y-%m-%d"
            			});
            			var CAL_End = new Calendar({
            				inputField : "eDt", trigger    : "eDt_trigger",
            				onSelect: function() {
            					var date = Calendar.intToDate(this.selection.get());
            					CAL_Start.args.max = date;
            					CAL_Start.redraw();
            					this.hide();
            				}, bottomBar: true, dateFormat: "%Y-%m-%d"
            			});
            		</script>
					</td>
				</tr>
				<tr id="iimgtr" style="display:<%=CHKIIF(pgprogbn="m" or pgprogbn="","none","")%>">
					<td bgcolor="#EFEFEF" width="100" align="center">�̹���</td>
					<td bgcolor="#FFFFFF">
					<input type="button" class="button" value="�̹������" onClick="jsSetImg('<%=cimage%>','cimage','spancimage')">
		   		    (��ٱ��Ͽ� ǥ�õǴ� �̹���)
		   		    <div id="spancimage" style="padding: 5 5 5 5">
		   		        <%IF cimage <> "" THEN %>
		   				<img  src="<%=cimage%>" />
		   				<a href="javascript:jsDelImg('cimage','spancimage');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">����</td>
					<td bgcolor="#FFFFFF">
                    <input type="text" class="text" size="20" name="conts" value="<%=conts%>" maxlength="30">
                    <br>(ex: 5�����̻� / 2,3���� )
                    <br>(ex: ����ī�� û������ 5% )
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">��ũ</td>
					<td bgcolor="#FFFFFF">
                    <input type="text" class="text" size="80" name="contlink" maxlength="100" value="<%=contlink%>">
                    <br>��ũ�� �ִ°�츸 �Է�
                    <br>(ex : http://www.10x10.co.kr/event/eventmain.asp?eventid=20960)
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">��뿩��</td>
				    <td bgcolor="#FFFFFF">
				    <input type="radio" name="isusing" value="Y" <%=CHKIIF(isusing="Y" or isusing="","checked","")%> >���
				    <input type="radio" name="isusing" value="N" <%=CHKIIF(isusing="N" ,"checked","")%> >������
				    </td>
				</tr>
				<% IF idx <> "" THEN%>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�����</td>
				    <td bgcolor="#FFFFFF">
				    <%=regdate%>
				    </td>
				</tr>
				<% end if %>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td align="left"><a href="javascript:self.close()"><img src="/images/icon_cancel.gif" border="0"></a></td>
					<td align="right"><input type="image" src="/images/icon_save.gif"></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->