using System.Collections.Generic;

public class TtsModel
{
    public string Id { get; set; }
    public string ModelType { get; set; }
    public string Developer { get; set; }
    public string Name { get; set; }
    public List<LanguageInfo> Language { get; set; } // Correctly defined as a list
    public string Quality { get; set; }
    public int SampleRate { get; set; }
    public int NumSpeakers { get; set; }
    public string Url { get; set; }
    public bool Compression { get; set; }
    public double FilesizeMb { get; set; }
}
