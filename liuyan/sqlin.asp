

<%
id=trim(request("id"))                                'ȡ��ID����ĺ���,Ϊ�����жϹؼ���.
if id="" or instr(id,"'")>0  or instr(id," ")>0 then  '���˵�SQL�����дʣ��ر��ǿո�� ��Ʋ��
    response.write "�����ύ����!"
    response.end
end if 
%>