
// Defines the structure for TTS model metadata.
public class TtsModel
{
    public string Id { get; set; }
    public string ModelType { get; set; }
    public string Developer { get; set; }
    public string Name { get; set; }
    public List<LanguageInfo> Language { get; set; }
    public string Url { get; set; }
}
