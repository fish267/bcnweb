

<!--#include file="conn.asp" -->
<!--#include file="md5.asp" -->
<%
if request.Form("pass")="" then                                    '�����ȡ��passΪ��,�������������.     
    response.Write"<script language=javascript>alert('δ������֤����');history.back(-1);</script>"
 response.End()                                                    '���ҽ�����������
 end if                                                            '����if���
 pass=md5(trim(request.form("pass")))                              '��passȥ���ո�,������md5��֤
 sql="select * from admin where pass='"&pass&"'"                   '��ѯ���ݿ�
 set rs=server.CreateObject("adodb.recordset")                     '����rs ����Ȩ��
 rs.open sql,conn,1,1                                              '�������ݿ�
 if rs.eof then                                                    '��������û�к������������ͬ����ת��err.html
      response.Redirect"err.html" 
	  response.End()                                               '����һ�в���
 else                                                              '����
      session("pass")=rs("pass")                                   'pass�������ݿ����pass��ת��index.asp 
      response.Redirect"index.asp"
 end if                                                            '����if ���
 rs.close                                                          '����һ��rs����
 set rs=nothing                                                     '���rs
%> 