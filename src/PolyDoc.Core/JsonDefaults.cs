using System.Text.Json;
using System.Text.Json.Serialization;

namespace PolyDoc.Core;

/// <summary>모든 코덱이 공유하는 기본 JsonSerializerOptions.</summary>
public static class JsonDefaults
{
    public static JsonSerializerOptions Options { get; } = Build();

    private static JsonSerializerOptions Build()
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DictionaryKeyPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true,
        };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
        options.Converters.Add(new BlockJsonConverter());
        options.Converters.Add(new FloatingObjectJsonConverter());
        return options;
    }
}
