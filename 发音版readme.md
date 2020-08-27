# 再要你命3000(发音版)
使用excel背单词的一个问题是缺少单词的发音，再要你命3000(发音版)能解决此问题。
## 使用说明
*提示：此版仅能在windows系统下正常使用。*
1. 下载3000_wav文件夹和3000(发音版/发音注释版).xlsm，放在同一个文件夹下。
2. 打开xlsm文件，并启用宏。
3. 点击单词的UK/US音标。
## 发音版简介
### wav音频文件
+ 命名格式 L(ist)_U(nit)_Word.wav List, Unit, Word对应列C、E、I
+ 源音频来自于[官方](http://www.dogwood.com.cn/mp3/yaoniming3000/)
+ 使用 [autosub](https://github.com/BingLingGroup/autosub) 对源音频进行分割
+ 使用google的语音识别识别每个单词，并人工校对。
### 宏
发音版和发音注释版所添加的宏完全一致
#### ReadWord
参考[How To Play A Sound If A Condition Is Met In Excel?](https://www.extendoffice.com/documents/excel/4417-excel-play-sound-if-condition-is-true.html)

    #If Win64 Then
        Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
            Alias "PlaySoundA" (ByVal lpszName As String, _
            ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean
    #Else
        Private Declare Function PlaySound Lib "winmm.dll" _
            Alias "PlaySoundA" (ByVal lpszName As String, _
            ByVal hModule As Long, ByVal dwFlags As Long) As Boolean
    #End If
        
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_FILENAME = &H20000
    Function ReadWord(list As String, unit As String, word As String) As String
        If Len(list) < 2 Then
            list = "0" + list
        End If
        Call PlaySound("3000_wav\" + list + "_" + unit + "_" + word + ".wav", _
          0, SND_ASYNC Or SND_FILENAME)
        SoundMe = ""
    End Function
#### onselect
    Option Explicit
     
    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        If Selection.Count = 1 Then
            Dim i As Range
            For Each i In Target
                If i.Row > 1 And (i.Column = 10 Or i.Column = 11) Then
                    Call ReadWord(Cells(i.Row, 3), Cells(i.Row, 5), Cells(i.Row, 9))
                End If
            Next
        End If
    End Sub

其中第7行判断了选择的范围（除第一行以外的10、11列，即单词的UK/US音标）。

第8行调用ReadWord，并传递当前单词的List, Unit, Word。




