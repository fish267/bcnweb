<!--
---------------------��    ��:reply.asp
---------------------��������:���԰����(��ȫע��)
---------------------��������:Dmlt
---------------------���߲���:Http://HI.Baidu.com/Int_yi
---------------------���߸���:�����BUG,����©д����,���뱾����ϵ;
---------------------��ϵ��ʽ:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<!--#include file="sqlin.asp" -->
<%
    if session("pass")="" then                                               '���passΪ��'�����һ������
	response.write "��û�е�½��֤��!"
	response.end
    end if                                                                   '����if ���
%>


<% 
dim rs3                                                                      '�������
dim sql3                                                                     '�������

id=request.form("id")                                                        '��ȡID

hf=request("hf")                                                             '��ȡhf
yd="1"                                                                       '��ֵ��yd

set rs3=server.CreateObject("adodb.recordset")                               '��rs3���ݿ����Ȩ��
sql3="select * from data where id="&cint(request.querystring("id"))          '��ѯ����
rs3.open sql3,conn,3,3                                                       '�������ݿ�
rs3("hf")=hf                                                                 '��ֵ��hf,���¸����ݿ�
rs3("yd")=yd                                                                 '��ֵ��yd,���¸����ݿ�
rs3.update                                                                   'rs3����Ϊupdate (���²���)
rs3.close                                                                    '����rs3һ�в���
response.write "<script>history.go(-1)</script>"                             '���ص���һҳ

%>