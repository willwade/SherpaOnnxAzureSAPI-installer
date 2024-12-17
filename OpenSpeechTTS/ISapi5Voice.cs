using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid(Guids.ISapi5VoiceInterfaceId)]
public interface ISapi5Voice
{
    void Speak(string text);
    void SetRate(int rate);
    void SetVolume(int volume);
    void Pause();
    void Resume();
}
