using System.Text.Json;
using System.Text.Json.Serialization;

namespace PolyDoc.Core;

/// <summary>
/// FloatingObject 다형 직렬화 컨버터.
/// 패턴은 <see cref="BlockJsonConverter"/> 와 동일 — discriminator <c>$type</c> 을 첫 속성으로 출력하고,
/// 읽기 시 누락되면 예외를 던진다 (FloatingObject 는 이번 사이클에 신규 도입이라 legacy 호환 없음).
/// </summary>
public sealed class FloatingObjectJsonConverter : JsonConverter<FloatingObject>
{
    public override FloatingObject Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartObject)
        {
            throw new JsonException($"Expected StartObject for FloatingObject, got {reader.TokenType}");
        }

        using var doc = JsonDocument.ParseValue(ref reader);
        var root = doc.RootElement;

        string? discriminator = null;
        if (root.TryGetProperty("$type", out var t) && t.ValueKind == JsonValueKind.String)
        {
            discriminator = t.GetString();
        }

        var targetType = discriminator switch
        {
            "textbox" => typeof(TextBoxObject),
            null      => throw new JsonException("FloatingObject requires '$type' discriminator."),
            _         => throw new JsonException($"Unknown FloatingObject discriminator: '{discriminator}'"),
        };

        var json = root.GetRawText();
        return (FloatingObject)JsonSerializer.Deserialize(json, targetType, options)!;
    }

    public override void Write(Utf8JsonWriter writer, FloatingObject value, JsonSerializerOptions options)
    {
        var type = value.GetType();
        var discriminator = type.Name switch
        {
            nameof(TextBoxObject) => "textbox",
            _ => throw new JsonException($"Unknown FloatingObject runtime type: {type.FullName}"),
        };

        var bytes = JsonSerializer.SerializeToUtf8Bytes(value, type, options);
        using var doc = JsonDocument.Parse(bytes);

        writer.WriteStartObject();
        writer.WriteString("$type", discriminator);
        foreach (var prop in doc.RootElement.EnumerateObject())
        {
            if (prop.NameEquals("$type")) continue;
            prop.WriteTo(writer);
        }
        writer.WriteEndObject();
    }
}
