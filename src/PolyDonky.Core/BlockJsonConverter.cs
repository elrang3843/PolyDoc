using System.Text.Json;
using System.Text.Json.Serialization;

namespace PolyDonky.Core;

/// <summary>
/// Block 다형 직렬화 컨버터.
///
/// 호환 정책:
/// - 읽기: <c>$type</c> 우선, 옛 빌드의 <c>kind</c> 도 허용. 둘 다 없으면 paragraph 로 폴백.
///   (29c09bd → c7b9de3 사이에 discriminator 가 "kind" → "$type" 으로 변경된 이력 때문.)
/// - 쓰기: 항상 <c>$type</c> 을 첫 속성으로 출력.
/// - 2026-04-29 통합: <c>textbox</c> discriminator 추가. 옛 FloatingObject 였던 글상자도
///   이제 Block 트리 안에서 함께 직렬화·역직렬화된다.
///
/// <see cref="System.Text.Json.Serialization.JsonPolymorphicAttribute"/> 의 기본 동작은 누락 discriminator 에
/// <c>NotSupportedException</c> 을 던지므로, 사용자의 옛 iwpf 파일을 무손실로 읽기 위해 커스텀.
/// </summary>
public sealed class BlockJsonConverter : JsonConverter<Block>
{
    public override Block Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartObject)
        {
            throw new JsonException($"Expected StartObject for Block, got {reader.TokenType}");
        }

        using var doc = JsonDocument.ParseValue(ref reader);
        var root = doc.RootElement;

        string? discriminator = null;
        if (root.TryGetProperty("$type", out var t1) && t1.ValueKind == JsonValueKind.String)
        {
            discriminator = t1.GetString();
        }
        else if (root.TryGetProperty("kind", out var t2) && t2.ValueKind == JsonValueKind.String)
        {
            discriminator = t2.GetString();
        }

        var targetType = discriminator switch
        {
            "paragraph" => typeof(Paragraph),
            "table"     => typeof(Table),
            "image"     => typeof(ImageBlock),
            "shape"     => typeof(ShapeObject),
            "textbox"   => typeof(TextBoxObject),
            "opaque"    => typeof(OpaqueBlock),
            null        => typeof(Paragraph),  // legacy fallback
            _           => throw new JsonException($"Unknown Block discriminator: '{discriminator}'"),
        };

        var json = root.GetRawText();
        return (Block)JsonSerializer.Deserialize(json, targetType, options)!;
    }

    public override void Write(Utf8JsonWriter writer, Block value, JsonSerializerOptions options)
    {
        var type = value.GetType();
        var discriminator = type.Name switch
        {
            nameof(Paragraph)     => "paragraph",
            nameof(Table)         => "table",
            nameof(ImageBlock)    => "image",
            nameof(ShapeObject)   => "shape",
            nameof(TextBoxObject) => "textbox",
            nameof(OpaqueBlock)   => "opaque",
            _ => throw new JsonException($"Unknown Block runtime type: {type.FullName}"),
        };

        // 구체 타입으로 직렬화 후 $type 을 맨 앞에 주입하여 다시 쓴다.
        var bytes = JsonSerializer.SerializeToUtf8Bytes(value, type, options);
        using var doc = JsonDocument.Parse(bytes);

        writer.WriteStartObject();
        writer.WriteString("$type", discriminator);
        foreach (var prop in doc.RootElement.EnumerateObject())
        {
            // "$type" 만 중복 방지. "kind" 는 ShapeObject.Kind / OpaqueBlock.Kind 등
            // 실제 도메인 속성 이름이므로 스킵하면 안 됨.
            // (읽기 경로에서 "$type" 이 "kind" 보다 우선하므로 레거시 파일 호환에도 문제 없음.)
            if (prop.NameEquals("$type")) continue;
            prop.WriteTo(writer);
        }
        writer.WriteEndObject();
    }
}
