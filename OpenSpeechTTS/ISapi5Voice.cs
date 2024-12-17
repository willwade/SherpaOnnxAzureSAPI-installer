using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("74c1c045-6e6a-4623-9735-2326a149af54")]
public interface ISapi5Voice
{
    void Speak(string text);
    void SetRate(int rate);
    void SetVolume(int volume);
    void Pause();
    void Resume();
}
