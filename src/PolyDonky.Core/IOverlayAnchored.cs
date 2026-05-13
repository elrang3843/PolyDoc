namespace PolyDonky.Core;

/// <summary>
/// 페이지 위에 자유 위치로 배치되는 모든 오버레이 객체(이미지·도형·표·글상자) 가
/// 구현하는 공통 anchoring 인터페이스.
///
/// IWPF.md 의 "개체 anchoring/배치 — 공통 모델 최대상한(superset)" 원칙에 따라
/// 좌표 체계를 모든 객체 타입에서 통일한다:
/// <list type="bullet">
///   <item><see cref="AnchorPageIndex"/> — 0-based 페이지 인덱스 (0 = 첫 페이지)</item>
///   <item><see cref="OverlayXMm"/> — 해당 페이지 좌상단 기준 X (mm)</item>
///   <item><see cref="OverlayYMm"/> — 해당 페이지 좌상단 기준 Y (mm)</item>
/// </list>
///
/// 좌표는 항상 *페이지 로컬* 이라 편집 화면과 인쇄 미리보기·인쇄 출력에서
/// 같은 의미로 해석된다 (이전 사이클의 "연속 Y" 좌표와 다름 — 마이그레이션은
/// <see cref="OverlayAnchorMigration"/> 참고).
/// </summary>
public interface IOverlayAnchored
{
    int AnchorPageIndex { get; set; }
    double OverlayXMm   { get; set; }
    double OverlayYMm   { get; set; }

    /// <summary>InFrontOfText/BehindText 모드에서의 페이지 로컬 X 좌표 (mm).
    /// OverlayXMm 의 별칭.</summary>
    double BlockX { get => OverlayXMm; set => OverlayXMm = value; }

    /// <summary>InFrontOfText/BehindText 모드에서의 페이지 로컬 Y 좌표 (mm).
    /// OverlayYMm 의 별칭.</summary>
    double BlockY { get => OverlayYMm; set => OverlayYMm = value; }
}
