<%

'������ȣ �ɼǹڽ�
Sub DrawInlinePhoneBox(byval extension)
	dim buf,i

	buf = "<select class='select' name='extension'>"
	if (""=CStr(extension)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
    for i=801 to 807
		if (CStr(i)=CStr(extension)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    for i=901 to 911
		if (CStr(i)=CStr(extension)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� �ð���
Sub DrawCallcenterHourBox(byval hour_from, hour_to)
	dim buf,i

	buf = "<select class='select' name='hour_from'>"
	if (""=CStr(hour_from)) then
		buf = buf + "<option value='' selected>--</option>"
	else
		buf = buf + "<option value='' >--</option>"
    end if
    for i=9 to 18
		if (CStr(i)=CStr(hour_from)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + " - "

	buf = buf + "<select class='select' name='hour_to'>"
	if (""=CStr(hour_to)) then
		buf = buf + "<option value='' selected>--</option>"
	else
		buf = buf + "<option value='' >--</option>"
    end if
    for i=9 to 18
		if (CStr(i)=CStr(hour_to)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� ���߽� �ɼǹڽ�
Sub DrawCallcenterInOutStateBox(byval dcontext)
	dim buf,i

	buf = "<select class='select' name='dcontext'>"
	if (""=CStr(dcontext)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("inbound"=CStr(dcontext)) then
		buf = buf + "<option value='inbound' selected>������ȭ</option>"
	else
		buf = buf + "<option value='inbound' >������ȭ</option>"
    end if
	if ("outbound"=CStr(dcontext)) then
		buf = buf + "<option value='outbound' selected>�߽���ȭ</option>"
	else
		buf = buf + "<option value='outbound' >�߽���ȭ</option>"
    end if
	if ("hunt_context"=CStr(dcontext)) then
		buf = buf + "<option value='hunt_context' selected>��Ʈ����</option>"
	else
		buf = buf + "<option value='hunt_context' >��Ʈ����</option>"
    end if
	if ("pers_context"=CStr(dcontext)) then
		buf = buf + "<option value='pers_context' selected>��������</option>"
	else
		buf = buf + "<option value='pers_context' >��������</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� ���� �ɼǹڽ�
Sub DrawCallcenterModeBox(byval mode)
	dim buf,i

	buf = "<select class='select' name='mode'>"
	if ("all"=CStr(mode)) then
		buf = buf + "<option value='all' selected>���κ���</option>"
	else
		buf = buf + "<option value='all' >���κ���</option>"
    end if
	
	if ("try"=CStr(mode)) then
		buf = buf + "<option value='try' selected>�õ���ȭ(��ü)</option>"
	else
		buf = buf + "<option value='try' >�õ���ȭ(��ü)</option>"
    end if
    
	if ("trycall"=CStr(mode)) then
		buf = buf + "<option value='trycall' selected>�õ���ȭ(����ü)</option>"
	else
		buf = buf + "<option value='trycall' >�õ���ȭ(����ü)</option>"
    end if
    
	if ("trycallnotplay"=CStr(mode)) then
		buf = buf + "<option value='trycallnotplay' selected>�õ���ȭ(��,�ٹ��ð�)</option>"
	else
		buf = buf + "<option value='trycallnotplay' >�õ���ȭ(��,�ٹ��ð�)</option>"
    end if
    
	if ("trycallonlyplay"=CStr(mode)) then
		buf = buf + "<option value='trycallonlyplay' selected>�õ���ȭ(��,�ٹ��ð���)</option>"
	else
		buf = buf + "<option value='trycallonlyplay' >�õ���ȭ(��,�ٹ��ð���)</option>"
    end if
    
	if ("successall"=CStr(mode)) then
		buf = buf + "<option value='successall' selected>������ȭ(��ü)</option>"
	else
		buf = buf + "<option value='successall' >������ȭ(��ü)</option>"
    end if
    
	if ("successcall"=CStr(mode)) then
		buf = buf + "<option value='successcall' selected>������ȭ(�ݼ���)</option>"
	else
		buf = buf + "<option value='successcall' >������ȭ(�ݼ���)</option>"
    end if
    
	if ("successnotcall"=CStr(mode)) then
		buf = buf + "<option value='successnotcall' selected>������ȭ(�ݼ�������)</option>"
	else
		buf = buf + "<option value='successnotcall' >������ȭ(�ݼ�������)</option>"
    end if
    
	if ("success2"=CStr(mode)) then
		buf = buf + "<option value='success2' selected>������ȭ(��������)</option>"
	else
		buf = buf + "<option value='success2' >������ȭ(��������)</option>"
    end if
    
	if ("outcall"=CStr(mode)) then
		buf = buf + "<option value='outcall' selected>�ݼ��͹߽���ȭ</option>"
	else
		buf = buf + "<option value='outcall' >�ݼ��͹߽���ȭ</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� �亯���� �ɼǹڽ�
Sub DrawCallcenterAnswerStateBox(byval dispositiono)
	dim buf,i

	buf = "<select class='select' name=dispositiono'>"
	if (""=CStr(dispositiono)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("ANSWERED9"=CStr(dispositiono)) then
		buf = buf + "<option value='ANSWERED' selected>ANSWERED</option>"
	else
		buf = buf + "<option value='ANSWERED' >ANSWERED</option>"
    end if
	if ("BUSY"=CStr(disposition)) then
		buf = buf + "<option value='BUSY' selected>BUSY</option>"
	else
		buf = buf + "<option value='BUSY' >BUSY</option>"
    end if
	if ("FAILED"=CStr(disposition)) then
		buf = buf + "<option value='FAILED' selected>FAILED</option>"
	else
		buf = buf + "<option value='FAILED' >FAILED</option>"
    end if
	if ("NO ANSWER"=CStr(disposition)) then
		buf = buf + "<option value='NO ANSWER' selected>NO ANSWER</option>"
	else
		buf = buf + "<option value='FAILED' >NO ANSWER</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� ��ȭ��ȣ �ɼǹڽ�
Sub DrawCallcenterPhoneNameBox(byval phoneno)
	dim buf,i

	buf = "<select class='select' name='phoneno'>"
	if (""=CStr(phoneno)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("07075490429"=CStr(phoneno)) then
		buf = buf + "<option value='07075490429' selected>�ݼ�����Ʈ</option>"
	else
		buf = buf + "<option value='07075490429' >�ݼ�����Ʈ</option>"
    end if
	if ("07075490556"=CStr(phoneno)) then
		buf = buf + "<option value='07075490556' selected>�繫����Ʈ</option>"
	else
		buf = buf + "<option value='07075490556' >�繫����Ʈ</option>"
    end if
	if ("07075490449"=CStr(phoneno)) then
		buf = buf + "<option value='07075490449' selected>��ǥ��ȣ2</option>"
	else
		buf = buf + "<option value='07075490449' >��ǥ��ȣ2</option>"
    end if
	if ("07075490448"=CStr(phoneno)) then
		buf = buf + "<option value='07075490448' selected>��ǥ��ȣ1</option>"
	else
		buf = buf + "<option value='07075490448' >��ǥ��ȣ1</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'�ݼ��� ���߽� ���ڿ�
Sub PrintCallcenterInOutState(byval dcontext)
	dim buf

	if ("inbound"=CStr(dcontext)) then
		buf = "����"
	elseif ("outbound"=CStr(dcontext)) then
    	buf = "�߽�"
	elseif ("toexten"=CStr(dcontext)) then
		'buf = "����"
		buf = CStr(dcontext)
    elseif ("hunt_context"=CStr(dcontext)) then
		buf = "��Ʈ"
    elseif ("pers_context"=CStr(dcontext)) then
		buf = "����"
    else
    	buf = CStr(dcontext)
    end if

    response.write buf
end Sub

'�ݼ��� ������ ���� ���ڿ�
Sub PrintCallcenterLastState(byval lastapp)
	dim buf

	if ("Playback"=CStr(lastapp)) then
		buf = "�ȳ���Ʈ"
	elseif ("Hangup"=CStr(lastapp)) then
    	buf = "��ȭ����"
	elseif ("Dial"=CStr(lastapp)) then
		buf = "��ȭ����"
	elseif ("BackGround"=CStr(lastapp)) then
		buf = "����Ʈ"
	elseif ("WaitExten"=CStr(lastapp)) then
		buf = "�������"
	elseif ("Busy"=CStr(lastapp)) then
		buf = "������"
	else
		buf = CStr(lastapp)
    end if

    response.write buf
end Sub

'�ݼ��� ��ȭ��ȣ ���ڿ�
Sub PrintCallcenterPhoneNumberString(byval dst)
	dim buf

	if ("07075490429"=CStr(dst)) then
		buf = "�ݼ�����Ʈ"
	elseif ("07075490556"=CStr(dst)) then
    	buf = "�繫����Ʈ"
	elseif ("Dial"=CStr(dst)) then
		buf = "��ȭ����"
	elseif ("07075490449"=CStr(dst)) then
		buf = "��ǥ��ȣ2"
	elseif ("07075490448"=CStr(dst)) then
		buf = "��ǥ��ȣ1"
	else
		buf = CStr(dst)
    end if

    response.write buf
end Sub

function SectoTime(v)
	dim temp, h, m, s
	if v > (60*60*24) then
		v = v mod 24
	end if
	h = int(v/3600)
	temp = v mod 3600
	m = int(temp/60)
	s = temp mod 60

	if (h < 10) then h = "0" & h
	if (m < 10) then m = "0" & m
	if (s < 10) then s = "0" & s

	sectotime = h &":"& m &":"& s
end function

%>