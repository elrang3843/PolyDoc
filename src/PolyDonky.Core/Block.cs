namespace PolyDonky.Core;

/// <summary>
/// 섹션 본문을 구성하는 블록 (문단·표·이미지 등) 의 추상 베이스.
/// 다형 직렬화는 <see cref="BlockJsonConverter"/> 에서 처리한다 (옛 'kind' discriminator 호환 위해).
/// </summary>
public abstract class Block
{
    public string? Id { get; set; }
    public NodeStatus Status { get; set; } = NodeStatus.Clean;

    /// <summary>블록의 너비 (mm). 0 이하면 자동/미정의.
    /// 치수를 직접 갖는 서브클래스(ImageBlock·ShapeObject·TextBoxObject·Table)에서는
    /// 해당 클래스의 WidthMm 을 반환하도록 override 된다.</summary>
    public virtual double BlockWidth  { get => 0; set { } }

    /// <summary>블록의 높이 (mm). 0 이하면 자동/미정의.
    /// 치수를 직접 갖는 서브클래스에서 override 된다.</summary>
    public virtual double BlockHeight { get => 0; set { } }
}

/// <summary>
/// Provenance / dirty tracking 상태.
/// 수정되지 않은 노드는 export 시 원본 조각을 재사용한다.
/// </summary>
public enum NodeStatus
{
    Clean,
    Modified,
    Opaque,
    Degraded,
}
