<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
If not (Request.ServerVariables("REMOTE_ADDR") = "61.252.133.75" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.105" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.106") Then
	dbget.close
	Response.End
End If

Response.CharSet = "euc-kr"

	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	Dim cEvtList, page, iPageSize, iPerCnt, isResearch, sDate, sSdate, sEdate, sEvt, strTxt, sKind, dispCate, strparm
	Dim arrList, iTotCnt, iTotalPage, sSort, blnReqPublish, intLoop, arreventkind, maxDepth, iStartPage, iEndPage, ix
	page = NullFillWith(requestCheckVar(Request("page"),10),1)	'���� ������ ��ȣ
	iPageSize = 10		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	maxDepth = 2
	'response.write page

	isResearch = NullFillWith(requestCheckVar(Request("isResearch"),1),"0")
	sDate 		= requestCheckVar(Request("selDate"),1)  	'�Ⱓ
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)
	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'�̺�Ʈ �ڵ�/�� �˻�
	strTxt 		= requestCheckVar(Request("sEtxt"),60)
	sKind 		= requestCheckVar(Request("eventkind"),32)	'�̺�Ʈ����
	dispCate 	= requestCheckvar(request("disp"),16)

	arreventkind= fnSetCommonCodeArr("eventkind",False)
	
	if isResearch="0" and sKind="" then
		skind="1,12,13,23,27,28,29,31"
	end if


	'�̺�Ʈ ù������ �����׸��� ���̵��� 
	IF (sKind="" and isResearch="0") or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD�μ���� (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�) - ���̷�(fotoark), ���ְ�(arlejk), ����ȭ(barbie8711) ����
			sKind = "1,12,13,16,17,23,24"
		else
			'��Ÿ (��������,��ü,��ǰ,�귣��,���̾,�׽���,�űԵ����̳�,�����,�귣��Week)
			sKind = "1,12,13,16,17,23,24,19,25,26,31"
		end if
	end if

	'#######################################
 	if sSort = "" then sSort = "CD"
 	if blnReqPublish= "" then blnReqPublish = False     
 	    
	'������ ��������
	set cEvtList = new ClsEvent
		cEvtList.FCPage = page		'����������
		cEvtList.FPSize = iPageSize		'���������� ���̴� ���ڵ尹��
		cEvtList.FSfDate 	= sDate		'�Ⱓ �˻� ����
		cEvtList.FSsDate 	= sSdate	'�˻� ������
		cEvtList.FSeDate 	= sEdate	'�˻� ������
		cEvtList.FSfEvt 	= sEvt		'�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt 	= strTxt	'�˻���
		cEvtList.FEDispCate	= dispCate	'�˻� ����ī�װ�
		cEvtList.FSkind 	= sKind
		cEvtList.FIsReqPublish = blnReqPublish
		cEvtList.FSort          = sSort
		
 		arrList = cEvtList.fnGetEventList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<form id="eventfrm" name="eventfrm" method="get" style="margin:0px;">
<input type="hidden" id="page" name="page">
<input type="hidden" name="isResearch" value="1"> 
<input type="hidden" name="sSort" value="<%=sSort%>">
<div class="searchWrap" style="border-top:none;">
	<div class="search">
		<ul>
			<li>
				<label class="formTit">�Ⱓ :</label>
				<select class="formSlt" id="selDate" name="selDate" title="�ɼ� ����">
			    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
			    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
			    	<option value="O" <%if Cstr(sDate) = "O" THEN %>selected<%END IF%>>������ ����</option>
				</select>
				<input type="text" class="formTxt" id="iSD" name="iSD" value="<%=sSdate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
				<img src="/images/admin_calendar.png" id="iSD_trigger" alt="�޷����� �˻�" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				~
				<input type="text" class="formTxt" id="iED" name="iED" value="<%=sEdate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
				<img src="/images/admin_calendar.png" id="iED_trigger" alt="�޷����� �˻�" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</li>
		</ul>
	</div>
	<dfn class="line"></dfn>
	<div class="search">
		<ul>
			<li>
				<p class="formTit">�̺�Ʈ ���� :</p>
				<%sbGetOptCommonCodeArr "eventkind", sKind, True,True,"onChange='javascript:document.frmEvt.submit();'"%>
			</li>
			<li>
				<p class="formTit">ī�װ� :</p>
				<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
			</li>
		</ul>
	</div>
	<dfn class="line"></dfn>
	<div class="search">
		<ul>
			<li>
				<label class="formTit" for="schWord">�˻��� :</label>
				<select class="formSlt" id="selEvt" name="selEvt">
					<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
					<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
					<option value="evt_tag" <%if Cstr(sEvt) = "evt_tag" THEN %>selected<%END IF%>>TAG</option>
					<option value="evt_sub" <%if Cstr(sEvt) = "evt_sub" THEN %>selected<%END IF%>>����ī��</option>
				</select>
				<input type="text" class="formTxt" id="sEtxt" name="sEtxt" value="<%=strTxt%>" maxlength="60" style="width:400px" onKeyUp="jsEventValueCheck();" onKeyPress="if (event.keyCode == 13){ NextPage(1,'event'); return false;}" />
			</li>
		</ul>
	</div>
	<input type="button" id="btnsearh1" class="schBtn" value="�˻�" onClick="NextPage(1,'event');" />
	<input type="button" id="btnsearh2" style="display:none;" class="schBtn" value="�˻�" onClick="alert('�˻����Դϴ�. ��ø� ��ٷ��ּ���.');" />
</div>
</form>
<div class="tbListWrap tMar15">
	<div class="rt pad10">
		<span>�˻���� : <strong><%=iTotCnt%></strong></span> <span class="lMar10">������ : <strong><%=page%> / <%=iTotalPage%></strong></span>
	</div>
	<ul class="thDataList">
		<li>
			<p class="cell05"></p>
			<p class="cell12">�̺�Ʈ �ڵ�</p>
			<p class="cell12">�̺�Ʈ ����</p>
			<p class="cell12">���</p>
			<p>�̺�Ʈ��</p>
			<p class="cell12">ī�װ�</p>
			<p class="cell12">������</p>
			<p class="cell12">������</p>
		</li>
	</ul>
	<ul class="tbDataList" id="contentslist">
    <%IF isArray(arrList) THEN

    	For intLoop = 0 To UBound(arrList,2)
    %>
		<li id="tr<%= arrList(0,intLoop) %>" style="cursor:pointer;">
			<p class="cell05"><input type="checkbox" name="contentsidx<%=arrList(0,intLoop)%>" id="contentsidx<%=arrList(0,intLoop)%>" value="<%=arrList(0,intLoop)%>" onClick="jsThisCheck('<%=arrList(0,intLoop)%>','event');" /></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(0,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><img src="<%=arrList(34,intLoop)%>" width="50" height="50" border="0" /></p>
			<p class="lt" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(4,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(26,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(5,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(6,intLoop)%></p>
		</li>
   <%	Next
   	END IF
   	
	iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1
	
	If (page mod iPerCnt) = 0 Then
		iEndPage = page
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
   %>
	</ul>
	<div class="ct tPad20 bPad20 cBk1">
		<% if (iStartPage-1 )> 0 then %><a href="javascript:NextPage(<%= iStartPage-1 %>,'event')">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(page) then
		%>
			<a href="javascript:NextPage(<%= ix %>,'event')"><span class="cRd1">[<%=ix%>]</span></a>
		<%		else %>
			<a href="javascript:NextPage(<%= ix %>,'event')">[<%=ix%>]</a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:NextPage(<%= ix %>,'event')">[next]</a>
		<% else %>[next]<% end if %>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->