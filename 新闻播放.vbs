'���ƣ����Ų���
'��;�����ںӱ���ˮ��ѧ���������Զ�����
'���ߣ��������루2015��ʯ���֣�QQ��510776295��
'ʹ�÷��������м���
'ע�⣺���ű���17�С�WScript.Sleep 5000��һ�䣬Ϊ�����ź��ӳ�5000����ģ�ⰴ�»س������ӳ���Ϊ�˱�֤���ſ�ʼ�����ٰ��س������е���ֵ�ɸ��ݼ�������������޸ġ�
On Error Resume Next

m = Month(Now)
d = Day(Now)
If Len(m) = 1 Then m = "0" & m
If Len(d) = 1 Then d = "0" & d
n = "\\Hzsvr00\���Ų���\����" & Year(Now) & m & d

Set WshShell = Wscript.CreateObject("Wscript.Shell")
WshShell.Run (n & ".rmvb")
WshShell.Run (n & ".mp4")
WScript.Sleep 5000
WshShell.SendKeys "{ENTER}"