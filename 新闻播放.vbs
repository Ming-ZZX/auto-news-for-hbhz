'名称：新闻播放
'用途：用于河北衡水中学的晚新闻自动播放
'作者：名正在想（2015级石佳乐，QQ：510776295）
'使用方法：运行即可
'注意：本脚本第17行“WScript.Sleep 5000”一句，为打开新闻后延迟5000毫秒模拟按下回车键，延迟是为了保证新闻开始播放再按回车，其中的数值可根据计算机性能自行修改。
On Error Resume Next

m = Month(Now)
d = Day(Now)
If Len(m) = 1 Then m = "0" & m
If Len(d) = 1 Then d = "0" & d
n = "\\Hzsvr00\新闻播放\新闻" & Year(Now) & m & d

Set WshShell = Wscript.CreateObject("Wscript.Shell")
WshShell.Run (n & ".rmvb")
WshShell.Run (n & ".mp4")
WScript.Sleep 5000
WshShell.SendKeys "{ENTER}"