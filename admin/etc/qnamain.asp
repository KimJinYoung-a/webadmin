<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->


<script>
function FlashEmbed(fid,fn,wd,ht,para,tranYn)
{
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="' + wd + '" height="' + ht + '" id="' + fid + '" align="middle">');
	document.write('<param name="allowScriptAccess" value="always">');
	document.write('<param name="movie" value="' + fn + para + '">');
	document.write('<param name="menu" value="false">');
	document.write('<param name="quality" value="high">');
	if(tranYn=='Y') {
		document.write('<param name="wmode" value="transparent">');}
	document.write('<embed src="' + fn + para + '" menu="false" quality="high" wmode="transparent" width="' + wd + '" height="' + ht + '" name="' + fid + '" align="middle" allowScriptAccess="always" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
	document.write('</object>');
}

function WMVEmbed(fid,fn,wd,ht)
{
	document.write('<object ID="' + fid + '" WIDTH="' + wd + '" HEIGHT="' + ht + '"  classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" CODEBASE=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab standby="Loading Microsoft?Windows? Media Player components..." type="application/x-oleobject">');
	document.write('<param name="Filename" value="' + fn + '">');
	document.write('<param name="AutoStart" value="true">');
	document.write('<param name="ShowControls" value="true">');
	document.write('<param name="ShowPositionControls" value="false">');
	document.write('<param name="ShowTracker" value="true">');
	document.write('<param name="ShowGotoBar" value="false">');
	document.write('<param name="ShowDisplay" value="false">');
	document.write('<param name="ShowStatusBar" value="true">');
	document.write('<embed type="application/x-mplayer2">');
	document.write('</object>');
}

//������ �˾� ���� 
function TnWinWmv(v,wd,ht){
	var popwin =window.open('http://www.10x10.co.kr/common/watch_wmv.asp?movie=' + v+'&wd='+wd+'&ht='+ht, 'wv', 'width=400,height=340,left=400,top=200,location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no');
    popwin.focus();
}

</script>

<script language='javascript'>//FlashEmbed('wg','http://www.inno.co.kr/flash/wg_teaser_mall.swf','','','','Y');</script>

<script language="javascript">//WMVEmbed('wv','http://okbuddy.co.kr/teachingpen.wmv','560','480','Y');</script>

<a href="javascript:TnWinWmv('http://okbuddy.co.kr/teachingpen.wmv','560','480');">�����󺸱�</a>
<% 
function getStr(var)
	if var<10 then
		getStr="0" + CStr(var)
	else
		getStr=CStr(var)
	end if
end Function
%>

<%
tmpDay =request("tmpDay")
if tmpDay ="" then tmpDay = now()

FirstDay = dateserial(year(tmpDay),month(tmpDay),1) ' �̹��� ù° ��
LastDay = dateadd("y",-1,dateadd("m",1,FirstDay)) '�̹��� ��������

totalDayCnt = datediff("Y",firstDay,LastDay) '�̴��� ����-1
FirstWeekDay = WeekDay(FirstDay)	'�̴�ù���� ����

%>
<%= now() %>
<script language="javascript" type="text/javascript">
function SubQnaDate(strDay){

	document.CalSubmitFrm.day.value= strDay;

	var strYear	 = document.CalSubmitFrm.year.value;
	var strMonth = document.CalSubmitFrm.mon.value;

	var today = new Date();
	var	strNYear	= today.getYear();
	var strNMonth	= today.getMonth() +1;
	var strNDay		= today.getDate()-1;

	var yn = confirm(strYear+ '��' + strMonth + '��' + strDay + '�� \n ~ ' + strNYear + '��' + strNMonth + '��' + strNDay + '��');

	if (yn){
		document.CalSubmitFrm.submit();
	}

}
</script>
<form name="CalSubmitFrm" method="post" action="qnamain_do.asp">
<input type="hidden" name="mode" value="qna">
<input type="hidden" name="year" value="<%= year(tmpDay)%>">
<input type="hidden" name="mon" value="<%= month(tmpDay)%>">
<input type="hidden" name="day" value="">
</form>
<table width="200" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr>
		<td align="center" colspan="7">Q&A������(excel)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center"><font size="2"><a href="?tmpDay=<%= dateadd("m",-1,tmpDay) %>">prev</a></font></td>
		<td colspan="3" align="center"><font size="2"><%= year(tmpDay) %>�� <%= MonthName(month(tmpDay),false) %></font></td>
	<td colspan="2" align="center"><font size="2"><a href="?tmpDay=<%= dateadd("m",1,tmpDay) %>">next</a></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><font color="red"><b>��</b></font></td>
		<td align="center"><b>��</b></td>
		<td align="center"><b>ȭ</b></td>
		<td align="center"><b>��</b></td>
		<td align="center"><b>��</b></td>
		<td align="center"><b>��</b></td>
		<td align="center"><font color="blue"><b>��</b></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">

		<% for i=1 to FirstWeekDay-1 %>
			<td align="center"></td>
		<% next %>

		<% for i=0 to TotalDayCnt %>
			<% if day(now())-1= i then %>
				<!-- ���ó�¥ �� -->
				<td align="center" bgcolor="#CCCCCC"><font color="#666666"><span onclick="SubQnaDate('<%= i+1 %>')" style="cursor:pointer"><%= i+1 %></span></font></td>
			<% else %>
				<% if (i + FirstWeekDay) mod 7 = 1  then %>
				<!-- �Ͽ��� -->
					<td align="center"><font color="#FF6666"><span onclick="SubQnaDate('<%= i+1 %>')" style="cursor:pointer"><%= i+1 %></span></font></td>
				<% elseif (i + FirstWeekDay) mod 7 =0  then %>
				<!-- ����� -->
					<td align="center"><font color="#6666FF"><span onclick="SubQnaDate('<%= i+1 %>')" style="cursor:pointer"><%= i+1 %></span></font></td>
				<% else %>
				<!-- ���� -->
					<td align="center"><font color="#666666"><span onclick="SubQnaDate('<%= i+1 %>')" style="cursor:pointer"><%= i+1 %></span></font></td>
				<% end if%>
			<% end if %>


			<% if (i + FirstWeekDay) mod 7 =0  then %>
		</tr>
		<tr bgcolor="#FFFFFF">
			<% end if %>

		<% next %>
	</tr>
</table>
<br /><br />
<table width="650" border="1" cellpadding="0" cellspacing="0" class="a">
	<form name="smsfrm" method="post" action="qnamain_do.asp">
	<input type="hidden" name="mode" value="sms">
	<tr>
		<td colspan="2" align="center"><b>���ں�����</b></td>
	</tr>
	<tr>
		<td width="120">������ ����</td>
		<td><input type="radio" name="inputmethod" value="hp">�ڵ���<input type="radio" name="inputmethod" value="userid" checked>���̵�</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="text" name="inputArray" value="" size="75"><br>
			�ڵ���:(xxx-xxxx-xxxx,zzz-zzzz-zzzz,....)�������� �Է�<br>
			���̵�:(aaaaa,bbbb,ccccc...)�������� �Է�
		</td>
	</tr>
	<tr>
		<td>�޽���</td>
		<td><textarea name="sendmsg" rows="4" cols="60"></textarea></td>
	</tr>
	<tr>
		<td>�߽��� ��ȭ��ȣ</td>
		<td><input type="text" name="sendnumber" size="13" value="1644-6030">(000-0000-0000)</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="������"></td>
	</tr>
	</form>
</table>
<a href="http://movie.10x10.co.kr/143962_Sound_List.xls">���Ϲޱ�</a>
<script>
function fnPlay(swfsrc){	
	ifrm = document.getElementById('Player');	
	obj= ifrm.contentWindow.document;
	
	obj.write('<object id="Player" classid="CLSID:22D6f312-B0F6-11D0-94AB-0080C74C7E95" codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701" width=0 height=0>');
	obj.write('<param name="AutoStart" value="True">');
	obj.write('<param name="TransparentAtStart" value="false">');
	obj.write('<param name="ShowControls" value="0">');
	obj.write('<param name="ShowDisplay" value="0">');
	obj.write('<param name="ShowStatusBar" value="0">');
	obj.write('<param name="AutoSize" value="0">');
	obj.write('<param name="AnimationAtStart" value="false">');
	obj.write('<param name="FileName" value="'+ swfsrc + '">');
	obj.write('</object>');
	
	obj.close();
}


function fnPlay22(swfsrc){	
	ifrm = document.getElementById('Player2');	
	ifrm.AutoStart=true;
}
function fnPlayout(){	
	ifrm = document.getElementById('Player2');	
	ifrm.AutoStart=false;
}



function fnPlay33(swfsrc){

	ifrm = document.getElementById('Player');	
	obj= ifrm.contentWindow.document;
	var str = "";

	str += "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width=0 height=0>";
	str += "<param name='movie' value='" +swfsrc+"'>";
	str += "<param name='quality' value=high>";
	str += "<embed src='" +swfsrc+"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=0 height=0>";
	str += "</embed> </object>";

 	obj.write(str);

}
</script>

<% 

Function regTest(sText)
	dim oReg
	
	set oReg= New RegExp
	
	oReg.Pattern  = "<[^>]*>"
	oReg.IgnoreCase = false
	oReg.Global = True
	regTest = oReg.Replace(sText,"")
	Set oReg = Nothing

End Function 

'response.write regTest(sTest)

dim sTest :sTest ="<p align='center'> <table width='600' border='0' align='center' cellspacing='0' cellpadding='0'> <tr> " &_
"<td align='center'> <img src='http://www.jsis.co.kr/skinwiz/viewer/market/aaaaa.jpg' border='0'> </td> " &_
"</tr> <tr> <td align='center'> <img src='http://www.jsis.co.kr/wizstock/080226_0211_09.jpg' border='0'><br><br> " &_
"<img src='http://www.jsis.co.kr/wizstock/080226_0212_09.jpg' border='0'><br><br> <img src='http://www.jsis.co.kr/wizstock/080226_0213_09.jpg' border='0'>" &_
"<br><br> <img src='http://www.jsis.co.kr/wizstock/080226_0216_09.jpg' border='0'><br><br> " &_
"<img src='http://www.jsis.co.kr/wizstock/080226_0226_09.jpg' border='0'><br><br> " &_
"<img src='http://www.jsis.co.kr/wizstock/080226_0229_09.jpg' border='0'><br><br> " &_
"<img src='http://www.jsis.co.kr/wizstock/080226_0231_09.jpg' border='0'><br><br> " &_
"<img src='http://www.jsis.co.kr/wizstock/080226_0225_09.jpg' border='0'><br><br> " &_
"<br><br> </td> </tr> <tr> <td align='center'> <br> <SPAN style='FONT-SIZE: 9pt'>" &_
"<FONT color=#666666> *�� ������*<br /> ���~�Ҹ�(�ܸ�)67cm �����ѷ�104cm ��ü����66cm<br /> " &_
"����:��ư(�긮)<br /> <br /> </FONT></SPAN> <br><br> </td> </tr> <tr> <td align='center'> " &_
"<img src='http://www.jsis.co.kr/skinwiz/viewer/market/wizwidform_03.jpg' border='0'> </td> </tr> </table> " &_
"</p>'"
'response.write regTest(sTest)
%>
<!--

<iframe name="Player" id="Player" src="" width="100" height="100"></iframe> 

<div onclick="fnPlay('http://movie.10x10.co.kr/101699_thankyou.wma');" onmouseout="fnPlay('');" style="width:100;height:20;border:1px solid #EDEDED">�Ҹ����1</div>
<div onclick="fnPlay('http://movie.10x10.co.kr/101694_congra.wma');" onmouseout="fnPlay('');" style="width:100;height:20;border:1px solid #EDEDED">�Ҹ����2</div>
<div onclick="fnPlay('http://movie.10x10.co.kr/101694_congra.wma');" onmouseout="fnPlay('');" style="width:100;height:20;border:1px solid #EDEDED">�Ҹ����3</div>

<script>//fnPlay('http://movie.10x10.co.kr/101699_thankyou.wma');</script>

<object id="Player2" classid="CLSID:22D6f312-B0F6-11D0-94AB-0080C74C7E95" codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701" width=0 height=0>
<param name="AutoStart" value="False">
<param name="TransparentAtStart" value="false">
<param name="ShowControls" value="0">
<param name="ShowDisplay" value="0">
<param name="ShowStatusBar" value="0">
<param name="AutoSize" value="0">
<param name="AnimationAtStart" value="false">
<param name="FileName" value="">
</object>
<iframe name="Player" id="Player" src="" width="100" height="100"></iframe><div align="center" style="cursor:pointer" onClick="fnPlay('http://movie.10x10.co.kr/101699_thankyou.wma');"><img src="http://fiximage.10x10.co.kr/web2007/common/sound-button.gif" width="153" height="47" /></div>
-->
<embed src='http://www.gagbag.co.kr/bbs/data/file/bbs4/563189150_ebb1c707_7Exmas-Jingle.swf' width='100' height='16'>
<%
'====================================================
'' �ӽ� sms ������
'====================================================
'dim temphp,tempuser,sql,msg,Pcnt,Ucnt
'tempuser="sengyun1,woori018,ncyber1004,zzini125,chzhrptkd,dwnara,kzones,satbuyl,iovelove44,spazio,lecher75,danmuzidal,roa02,ksh7035,lovejuok"
'temphp="016-372-3924,011-9822-7357,019-313-6680,016-792-7952,011-867-6933,018-677-0468,011-9961-4146,010-4633-6224,016-296-0442,011-9048-0517,010-3013-0013,016-230-4063,016-278-8031,017-874-2348,016-423-4748,011-399-3091,016-780-9402,011-9719-7561,011-9907-3548,010-3080-8356,016-605-1101,011-9447-0912,011-9204-5389,011-9278-4967,010-3929-4510,011-9169-9902,016-9344-4223,011-757-8067,010-6324-8492"
'temphp=trim(temphp)


'temphp=split(temphp,",")
'Pcnt=ubound(temphp)

'tempuser=split(tempuser,",")
'Ucnt=ubound(tempuser)

'msg="�� ���Ŭ�� ������ ��÷! ���� Ȯ�����ּ���~[�ٹ�����]"

'for i=0 to Pcnt
'sql = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg )" + vbcrlf
'sql = sql + "values(" + vbcrlf
'sql = sql + "'" + temphp(i) +"'," + vbcrlf
'sql = sql + "'010-9979-0522','1',getdate(),'" + msg + "')" + vbcrlf

'rsget.open sql,dbget,1
'response.write sql & "(" & i & ")" & " �����Ͽ����ϴ�.<br>"
'next

'for i=0 to 48
'sql = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg )" + vbcrlf
'sql = sql + "values(" + vbcrlf
'sql = sql + "'" + temphp(i) +"'," + vbcrlf
'sql = sql + "'02-554-2052','1',getdate(),'" + msg2 + "')" + vbcrlf

'rsget.open sql,dbget,1
'response.write sql & " �����Ͽ����ϴ�.<br>"
'next

'====================================================
'' �ŴϾ� ��÷�� �Է�
'====================================================

'dim tempname,tempmail
'dim newarray(50)
'dim tempid
'tempid="alavua,alsnzjffj,astrud,bamandog,biveree,bluejiyong02,bseo,chueul,chyopaa,coolmvp1,crys8318,djstjd,doona0723,dorothy1007,dreams99,eastsky,eom0827,find12,geotec,greatoy,h2y0314,helle79,hyosun2,i486jesus,jaymom,jikka1017,jjj0519,jsmitwo,juhee,kaknh,keki,khjsky25,kimseri,kimssov,kohito,komicoco,lainep06,lgx19s,libran1003,LoLo7,miffy2486,mincloud,mliorvaen,mynamesopy,mysticu,nazimk,neko0930,nikino,nymph007,parakiss,piano3574,pinkhachi,polohjh,pooh6242,ragy,raksha,realrnjs,redrain27,ringfoz,sadmars,screamq,sharara1,shinanze,shinhan,shinycandy,sinsiyoo,siridoctor,sirosa,snabs2000,socho12,ssamttang,styulina,suna207,waeby,whome80,wild7168,worlygun,yhj5535,ylem3010,zzz120"
'tempid=split(tempid,",")
'tempname="������,������,�̾Ƹ�,�̿���,���ֿ�,��ϼ�,�ڿ���,�̼���,�����,������,�����,���̸�,���°�,������,�ڼ���,�赵��,��ȭ��,�輱��,������,����ȭ,�迵��,���ϳ�,�Ǻ���,������,����,������,�����,���ֿ�,Ȳ����,������,�̱Ϳ�,���ʷ�,�̿���,������,����,������,�輼��,������,����,����,�����,�Ѽ���,������,������,������,������,���¿�,��ҿ�,������,��ȿ��"
'tempname=split(tempname,",")
'tempmail="000706@hanmail.net,5ho5ho@hanmail.net,lal77@freechal.com,oops79@entaz.com,corncandy83@hanmail.net,mysteryls80@hanmail.net,a4953@hanmail.net,leesunhye@hotmail.com,yobebedh@hotmail.com,u-zin@hanmail.net,soogoo9@yahoo.co.kr,cristalle81@naver.com,d\0x5Fangelo@nate.com,ddubidduba@hotmail.com,dooky81@naver.com,ehrns2@paran.com,eureca2@nate.com,finetree97@hotmail.com,nervousfish@hanmail.net,gisel21witch@hotmail.com,haeter@hanmail.net,hn1111@nate.com,inspire77@korea.com,angelmin22@hanmail.net,jjangaya33@nate.com,bluebell@empal.com,joonie81@hanafos.com,joy2365@nate.com,lush2000@hanmail.net,lim_eun_ha@hotmail.com,se24@nate.com,thdchfhd@hanmail.net,blu-pepe@hanmail.net,hd1123.cho@samsung.com,letitbe2002@naver.com,nanotbo@hotmail.com,okokida@hanmail.net,vanness4@naver.com,hyuky@gshs.co.kr,0105kkr@naver.com,poporito@nate.com,soonju80@hanmail.net,sstary@daum.net,kyy214@nate.com,seryuni@hanmail.net,gjguswls@hotmail.com,silvia82@empal.com,sweeteggroll@naver.com,sommus@hanmail.net,loveivette@hotmail.com"
'tempmail=split(tempmail,",")
'for i=0 to 79
'	dim sql

'	sql = "insert into [db_contents].[dbo].tbl_mania_user(yyyymm,userid,point,coupon,gubun)" + vbcrlf
'	sql = sql + " values('2006-04','" + CStr(tempid(i)) + "',0,0,'07')"+vbcrlf
'	response.write i & "<br>"
'	response.write sql
'	rsget.open sql,dbget,1

'next


%>
<%'
'dim sql ,userid
'sql = " select userid from db_cts.dbo.tbl_ngene_event " &_
'			" where gubun='16' "
'
'db2_rsget.open sql ,db2_dbget,1
'
'if not db2_rsget.eof then
'	do until db2_rsget.eof
'		userid = userid & db2_rsget("userid") & ","
'	db2_rsget.movenext
'loop
'end if
'db2_rsget.close

'response.write userid
%>


<%'
'dim tempsql
'tempsql= "select comidx " &_
'					",Case SubmasterIdx  " &_
'					"	when 1 then 63373 " &_
'					"	when 2 then 36735 " &_
'					"	when 3 then 70647 " &_
'					"	when 4 then 65714 " &_
'					"	when 5 then 76544 " &_
'					"End as itemid " &_
'					",Case SubmasterIdx " &_
'					"	when 1 then '�����̺�' " &_
'					"	when 2 then 'Ź��ð�' " &_
'					"	when 3 then 'FishEye' " &_
'					"	when 4 then '��ũ��ȭ��' " &_
'					"	when 5 then '�𽺴� ���Ĺ' " &_
'					"End as itemname " &_
'					",Comuserid " &_
'					",comContents " &_
'					",t.ss " &_
'					"from db_cts.dbo.tbl_etcevent_comment c " &_
'					"Left join  " &_
'					"	(select userid,sum(totalcost) as ss  " &_
'					"	from [110.93.128.72].[db_order].[dbo].tbl_order_master " &_
'					"	where userid<>'' " &_
'					"	and ipkumdiv>=4 and ipkumdiv<9 " &_
'					"	and cancelyn='N' " &_
'					"	group by userid) as t " &_
'					"	on c.Comuserid= t.userid " &_
'					"where C.MasterEventIdx='4' " &_
'					"and ComIsusing='Y' "

'db2_rsget.open tempsql,db2_dbget,1


%>
<!--
<table border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>���̵� </td>
		<td>��ǰ �ڵ�</td>
		<td>��ǰ��</td>
		<td>����</td>
		<td>���ž�(6������)</td>

	</tr>
	<%' if not db2_rsget.eof then %>
	<%' do until db2_rsget.eof %>
	<tr>
		<td><%'= db2_rsget("Comuserid") %></td>
		<td><%'= db2_rsget("itemid") %></td>
		<td><%'= db2html(db2_rsget("itemname")) %></td>
		<td><%'= db2html(db2_rsget("ComContents")) %></td>
		<td><%'= db2html(db2_rsget("ss")) %></td>
	</tr>
	<%' db2_rsget.MoveNext %>
	<%' loop %>
	<%' end if%>
</table>
-->
<%
''dim sql
'sql ="" &_
'"select " &_
'"n.userid,n.username,n.juminno " &_
'",isnull(t.ss,0) as sm " &_
'",r.itemid,i.itemname,i.sellcash " &_
'"from [db_contents].[dbo].tbl_recommend_item r " &_
'"JOIN db_item.[dbo].tbl_item i " &_
'"	ON r.itemid=i.itemid " &_
'"JOIN db_user.[dbo].tbl_user_n n " &_
'"	ON r.userid=n.userid " &_
'"LEFT JOIN( " &_
'"	select m.userid,sum(d.itemcost) as ss " &_
'"	from db_order.[dbo].tbl_order_master m " &_
'"	JOIN db_order.[dbo].tbl_order_detail d " &_
'"		ON m.orderserial = d.orderserial and d.itemid<>0 " &_
'"	WHERE m.cancelyn='N'  " &_
'"		and d.cancelyn<>'Y'  " &_
'"		and m.jumundiv<>9 " &_
'"		and m.accountdiv<>'30' " &_
'"		and m.userid<>'' " &_
'"	group by m.userid " &_
'"	) as t " &_
'"	ON t.userid=r.userid " &_
'"where r.isusing='Y' " &_
'"and r.gubun=4 " &_
'"order by r.idx desc "

'rsget.open sql,dbget,1

'if not rsget.eof then
%>
	<table border='1' cellpadding='0' cellspacing='0'>
		<tr>
			<td>���̵�</td>
			<td>�̸�</td>
			<td>�ֹι�ȣ</td>
			<td>���ų���</td>
			<td>��ǰ��ȣ</td>
			<td>��ǰ��</td>
			<td>��ǰ ����</td>
		</tr>
	<%' do until rsget.eof %>
		<tr>
			<td><%'= rsget("userid") %></td>
			<td><%'= rsget("username") %></td>
			<td><%'= rsget("juminno") %></td>
			<td><%'= rsget("sm") %></td>
			<td><%'= rsget("itemid") %></td>
			<td><%'= db2html(rsget("itemname")) %></td>
			<td><%'= rsget("sellcash") %></td>
		</tr>

	<%' rsget.movenext
	'loop %>
	</table>
<%' end if %>
<%' rsget.close %>
<%' db2_rsget.close %>
<%
dim  ItemHTML_Basic
 ItemHTML_Basic ="" &_
"<tr> " &_
"	<td> " &_
"		<table width=""548"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd""> " &_
"			<tr> " &_
"				<td width=""260"" align=""right"" style=""border-right: 1px solid #dddddd""> " &_
"					<table width=""255"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0""> " &_
"						<tr> " &_
"							<td width=""50"" valign=""bottom""> " &_
"								<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0""> " &_
"									<tr> " &_
"										<td><img src=""::ITEMICONIMAGE::"" width=""50"" height=""50""></td> " &_
"									</tr> " &_
"									<tr> " &_
"										<td height=""17"" align=""center"" valign=""bottom"">::ITEMID::</td> " &_
"									</tr> " &_
"								</table> " &_
"							</td> " &_
"							<td  style=""padding:5"">::ITEMNAME::<br>[ ::ITEMOPTOINNAME:: ]</td> " &_
"						</tr> " &_
"					</table> " &_
"				</td> " &_
"				<td align=""center""> " &_
"					<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#eeeeee""> " &_
"						<tr> " &_
"							<td width=""60"" height=""35"" align=""center"">�� ��</td> " &_
"							<td width=""40"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">::ITEMNO::</td> " &_
"							<td width=""60"" align=""center"" style=""padding:0 5 0 5;"">�����Ȳ</td> " &_
"							<td style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">::DELIVERYSTATS::</td> " &_
"						</tr> " &_
"						<tr height=""1""> " &_
"							<td colspan=""4"" align=""center"" bgcolor=""#dddddd""></td> " &_
"						</tr> " &_
"						<tr> " &_
"							<td align=""center"">�����</td> " &_
"							<td colspan=""3"" style=""padding:5"" bgcolor=""#FFFFFF""><strong class=""Information_font"">::DELIVERYLINKTXT::</strong></td> " &_
"						</tr> " &_
"					</table> " &_
"				</td> " &_
"			</tr> " &_
"		</table> " &_
"	</td> " &_
"</tr> " 

'response.write ItemHTML_Basic
%>
<!-- include virtual="/lib/db/db2close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->