Attribute VB_Name = "modOptionalGlobal"
Public m_iActiveAlertWindows As Long

Private m_snd() As Byte
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file

Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
                                           (lpData As Any, _
                                      ByVal hModule As Long, _
                                      ByVal dwFlags As Long) As Long

Public Function PlaySound(ByVal SndID As Long) As Long
       Const Flags = SND_ASYNC Or SND_MEMORY
       m_snd = LoadResData(SndID, "SOUND")
       PlaySoundData m_snd(0), 0, Flags
End Function



