<!--
---------------------��    ��:add.asp
---------------------��������:���԰����(��ȫע��)
---------------------��������:Dmlt
---------------------���߲���:Http://HI.Baidu.com/Int_yi
---------------------���߸���:�����BUG,����©д����,���뱾����ϵ;
---------------------��ϵ��ʽ:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<%
name=trim(request("name"))                                                 '��ȡname,����ȥ������Ŀո�(trim ������ȥ�ո�)
if name=""  then                                                           '���name��ֵΪ��,�����һ������
response.write "<script>alert('û���ǳ�');location='index.asp';onclick=history.go(-1)</script>"
    response.end                                                           '���ҽ�����������
end if 
ly=trim(request("ly"))                                                    '��ȡly,����ȥ������Ŀո�(trim ������ȥ�ո�)
if ly=""  then                                                            '���ly��ֵΪ��,�����һ������
response.write "<script>alert('û������');location='index.asp';onclick=history.go(-1)</script>"
    response.end                                                          '���ҽ�����������
end if                                                                    '����if ���
%>

<%
set rs=server.createobject("adodb.recordset")                              '��rs����Ȩ��
sql="select * from data"                                                   '�Ա�Ĳ���
rs.open sql,conn,3,3                                                       '�������ݿ�
rs.addnew                                                                  '���������
rs("ly")=request.form("ly")                                                '��ֵ��ly,�����ݿ�
rs("name")=request.form("name")                                            '��ֵ��name,�����ݿ�
rs.update                                                                  'rsΪupdate
rs.close                                                                   '����һ��rs����
set rs=nothing                                                             '���rs
conn.close                                                                 '����һ��conn����
set conn=nothing                                                           '���conn
response.write "<script>history.go(-1)</script>"                           '������һҳ
%>