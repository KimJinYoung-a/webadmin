<%
Dim infoLoop, infoDivValue
Dim foreignMall : foreignMall = "N"
Dim hiddenMall : hiddenMall = "N"
Dim mLoop, vPurchasetype
vPurchasetype = request("purchasetype")
Select Case Request.ServerVariables("SCRIPT_NAME")
	Case "/admin/etc/my11st/my11stItem.asp"				foreignMall = "Y"
	Case "/admin/etc/zilingo/zilingoItem.asp"			foreignMall = "Y"
	Case "/admin/etc/shopify/shopifyItem.asp"			foreignMall = "Y"
	Case "/admin/etc/shopify/shopifyNewItem.asp"		foreignMall = "Y"
	Case "/admin/etc/nvstorefarmClass/nvClassItem.asp"	hiddenMall = "Y"
End Select
%>
<script language='javascript'>
function checkComp(comp){
	if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
		if ((comp.name=="bestOrd")&&(comp.checked)){
			comp.form.bestOrdMall.checked=false;
		}
		if ((comp.name=="bestOrdMall")&&(comp.checked)){
			comp.form.bestOrd.checked=false;
		}
	}
}
</script>
<label><input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(10x10)</b></label>&nbsp;
<label><input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(���޸�)</b></label>&nbsp;
<br />
�Ǹ�(10x10)
<select name="sellyn" class="select">
	<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
	<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
</select>&nbsp;
����
<select name="limityn" class="select">
	<option value="">��ü
	<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
	<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
</select>&nbsp;
����
<select name="sailyn" class="select">
	<option value="">��ü
	<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
	<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
</select>&nbsp;
<% If CMALLNAME = "11st1010" OR CMALLNAME = "lfmall" OR CMALLNAME = "cjmall" OR CMALLNAME = "interpark" OR CMALLNAME = "shintvshopping" OR CMALLNAME = "skstoa" OR CMALLNAME = "wetoo1300k" OR CMALLNAME = "lotteimall" OR CMALLNAME = "kakaostore" OR CMALLNAME = "boribori1010" OR CMALLNAME = "wconcept1010" OR CMALLNAME = "benepia1010" OR CMALLNAME = "qooi1010" OR CMALLNAME = "gmarket1010" OR CMALLNAME = "auction1010" Then %>
����(<%= getOutmallstandardMargin %>%)
<% Else %>
����(<%= CMAXMARGIN %>%)
<% End If %>
<select name="startMargin" class="select">
	<option value="">-����-</option>
	<% For mLoop = 0 to 100 %>
	<option value="<%= mLoop %>" <%= CHkIIF(CStr(startMargin) = CStr(mLoop),"selected","") %>><%= mLoop %></option>
	<% Next %>
</select>
~
<select name="endMargin" class="select">
	<option value="">-����-</option>
	<% For mLoop = 0 to 100 %>
	<option value="<%= mLoop %>" <%= CHkIIF(CStr(endMargin) = CStr(mLoop),"selected","") %>><%= mLoop %></option>
	<% Next %>
</select>&nbsp;
<% If hiddenMall = "N" Then %>
����
<select name="isMadeHand" class="select">
	<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü</option>
	<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y</option>
	<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N</option>
	<option value="T" <%= CHkIIF(isMadeHand="T","selected","") %> >�ֹ����۹���</option>
</select>&nbsp;
<% End If %>
�ɼ�
<select name="isOption" class="select">
	<option value="" <%= CHkIIF(isOption="","selected","") %> >��ü
	<option value="optAll" <%= CHkIIF(isOption="optAll","selected","") %> >�ɼ���ü
	<option value="optaddpricey" <%= CHkIIF(isOption="optaddpricey","selected","") %> >�߰��ݾ�Y
	<option value="optaddpricen" <%= CHkIIF(isOption="optaddpricen","selected","") %> >�߰��ݾ�N
	<option value="optN" <%= CHkIIF(isOption="optN","selected","") %> >��ǰ
</select>&nbsp;
ǰ��
<select name="infodiv" class="select">
	<option value="" <%= CHkIIF(infoDiv="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >�Է�
	<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >���Է�
<%
	For infoLoop = 1 To 35
		If infoLoop < 10 Then
			infoDivValue = "0"&infoLoop
		Else
			infoDivValue = infoLoop
		End If
%>
	<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
<% Next %>
	<option value="47" <%= CHkIIF(CStr(infodiv) = "47","selected","") %> >47
	<option value="48" <%= CHkIIF(CStr(infodiv) = "48","selected","") %> >48
</select>&nbsp;
<% If foreignMall = "N" Then %>
	<% If hiddenMall = "N" Then %>
���ܺ귣��
<select name="notinmakerid" class="select">
	<option value="" <%= CHkIIF(notinmakerid="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(notinmakerid="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(notinmakerid="N","selected","") %> >N
</select>&nbsp;
	<% End If %>
���ܻ�ǰ
<select name="notinitemid" class="select">
	<option value="" <%= CHkIIF(notinitemid="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(notinitemid="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(notinitemid="N","selected","") %> >N
</select>&nbsp;
�ɼ��߰��ݾ�
<select name="priceOption" class="select">
	<option value="" <%= CHkIIF(priceOption="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(priceOption="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(priceOption="N","selected","") %> >N
</select>&nbsp;
	<% If hiddenMall = "N" Then %>
Ư��
<select name="isSpecialPrice" class="select">
    <option value="" <%= CHkIIF(isSpecialPrice="","selected","") %> >��ü
    <option value="Y" <%= CHkIIF(isSpecialPrice="Y","selected","") %> >Y
</select>&nbsp;
	<% End If %>
<% End If %>
<br />
�Ǹ�(���޸�)
<select name="extsellyn" class="select">
	<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
	<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
<% If cmallname ="gsshop" Then %>
	<option value="E" <%= CHkIIF(extsellyn="E","selected","") %> >���
<% Else %>
	<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
	<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
	<% If cmallname ="interpark" Then %>
	<option value="SP" <%= CHkIIF(extsellyn="SP","selected","") %> >���Է�
	<% End If %>
<% End If %>
</select>&nbsp;
�������ܻ�ǰ
<select name="exctrans" class="select">
	<option value="" <%= CHkIIF(exctrans="","selected","") %> >��ü</option>
	<option value="Y" <%= CHkIIF(exctrans="Y","selected","") %> >Y</option>
	<option value="N" <%= CHkIIF(exctrans="N","selected","") %> >N</option>
	<option value="F" <%= CHkIIF(exctrans="F","selected","") %> >N(FAIL)</option>
</select>&nbsp;
����
<select name="failCntExists" class="select">
	<option value="" <%= CHkIIF(failCntExists="","selected","") %> >��ü</option>
	<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >��ϼ�������1ȸ�̻�</option>
	<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >��ϼ�������0ȸ</option>
	<option value="5U" <%= CHkIIF(failCntExists="5U","selected","") %> >����6ȸ �̻�</option>
	<option value="5D" <%= CHkIIF(failCntExists="5D","selected","") %> >����5ȸ ����</option>
</select>&nbsp;
��۱���
<% drawBeadalDiv "deliverytype", deliverytype %>&nbsp;
�ŷ�����
<% drawSelectBoxMWU "mwdiv", mwdiv %>
<% If CMALLNAME = "coupang" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	�����ġ
	<select name="GosiEqual" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(GosiEqual="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(GosiEqual="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	�����
	<select name="MatchShipping" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchShipping="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchShipping="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	�ɼǼ����� :
	<select name="regedOptOver" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(regedOptOver="Y","selected","") %> >�ʰ�
		<option value="N" <%= CHkIIF(regedOptOver="N","selected","") %> >�̸�
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "ssg" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���븶��
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="2" maxlength="2">
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "wetoo1300k" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	�귣��
	<select name="MatchBrand" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchBrand="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchBrand="N","selected","") %> >�̸�Ī
	</select>&nbsp;
<% ElseIf CMALLNAME = "hmall1010" Then %>
	�̹���
	<select name="MatchIMG" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchIMG="Y","selected","") %> >���
		<option value="N" <%= CHkIIF(MatchIMG="N","selected","") %> >�̵��
	</select>&nbsp;
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
	���븶��
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="5" maxlength="5">
<% ElseIf CMALLNAME = "auction1010" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "ezwel" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	��ǰ�з�
	<select name="MatchPrddiv" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
	</select>&nbsp;
<% ElseIf CMALLNAME = "gmarket1010" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
	G9��Ͽ���
	<select name="MatchG9" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchG9="Y","selected","") %> >���
		<option value="N" <%= CHkIIF(MatchG9="N","selected","") %> >�̵��
	</select>&nbsp;
	�ݾ�
	<select name="sellpriceChk" class="select">
		<option value="">��ü
		<option value="samman" <%= CHkIIF(sellpriceChk="samman","selected","") %> >3�����̻�
	</select>&nbsp;
<% ElseIf CMALLNAME = "gsshop" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	��ǰ�з�
	<select name="MatchPrddiv" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "nvstorefarm" or CMALLNAME = "nvstoregift" or CMALLNAME = "Mylittlewhoopee" or CMALLNAME = "WMP" or CMALLNAME = "interpark" or CMALLNAME = "lfmall" or CMALLNAME = "11st1010" or CMALLNAME = "shintvshopping" or CMALLNAME = "skstoa" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	<% If CMALLNAME = "lfmall" Then %>
	ǰ��з�
	<select name="MatchDiv" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchDiv="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchDiv="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	<% End If %>
	<% If CMALLNAME = "skstoa" Then %>
	���븶��
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="2" maxlength="2">
	<% End If %>
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "lotteon" or CMALLNAME = "lotteimall" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "cjmall" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	��ǰ�з�
	<select name="MatchPrddiv" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
	</select>&nbsp;
<% ElseIf CMALLNAME = "boribori1010" OR CMALLNAME = "wconcept1010" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
	�귣��
	<select name="MatchBrand" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchBrand="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchBrand="N","selected","") %> >�̸�Ī
	</select>&nbsp;
<% ElseIf CMALLNAME = "qooi1010" OR CMALLNAME = "benepia1010" Then %>
	ī�װ�
	<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
	</select>&nbsp;
<% ElseIf CMALLNAME = "sabangnet" Then %>
	���������ܻ�ǰ
	<select name="scheduleNotInItemid" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% End If %>
<br />
���޻��(��ǰ)
<select name="isextusing" class="select">
	<option value="">��ü</option>
	<option value="Y" <%= CHkIIF(isextusing="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(isextusing="N","selected","") %> >N
</select>&nbsp;
���޻��(�귣��)
<select name="cisextusing" class="select">
	<option value="">��ü</option>
	<option value="Y" <%= CHkIIF(cisextusing="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(cisextusing="N","selected","") %> >N
</select>&nbsp;
3�����Ǹŷ�
<select name="rctsellcnt" class="select">
	<option value="">��ü</option>
	<option value="0" <%= CHkIIF(rctsellcnt="0","selected","") %> >0
	<option value="1" <%= CHkIIF(rctsellcnt="1","selected","") %> >1���̻�
</select>&nbsp;
��������: 
<select name="purchasetype" class="select">
	<option value="">��ü</option>
	<option value="1"	<%= CHkIIF(vPurchasetype="1","selected","") %> >�Ϲ�����
	<option value="3"	<%= CHkIIF(vPurchasetype="3","selected","") %> >PB
	<option value="4"	<%= CHkIIF(vPurchasetype="4","selected","") %> >����
	<option value="5"	<%= CHkIIF(vPurchasetype="5","selected","") %> >ODM
	<option value="7"	<%= CHkIIF(vPurchasetype="7","selected","") %> >�귣�����
	<option value="6"	<%= CHkIIF(vPurchasetype="6","selected","") %> >����
	<option value="8"	<%= CHkIIF(vPurchasetype="8","selected","") %> >����
	<option value="9"	<%= CHkIIF(vPurchasetype="9","selected","") %> >�ؿ�����
	<option value="10"	<%= CHkIIF(vPurchasetype="10","selected","") %> >B2B
	<option value="356"	<%= CHkIIF(vPurchasetype="356","selected","") %> >PB/ODM/���Ը�
	<option value="101"	<%= CHkIIF(vPurchasetype="101","selected","") %> >�Ϲ����� ����
	<option value="102"	<%= CHkIIF(vPurchasetype="102","selected","") %> >������ǰ��
</select>&nbsp;