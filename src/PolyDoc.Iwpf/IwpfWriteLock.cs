namespace PolyDoc.Iwpf;

/// <summary>IWPF 암호 보호 모드.</summary>
public enum PasswordMode
{
    /// <summary>암호 보호 없음.</summary>
    None  = 0,

    /// <summary>열기 보호 — AES-256-GCM 으로 파일 전체를 암호화.</summary>
    Read  = 1,

    /// <summary>쓰기 보호 — 저장 시 비밀번호를 PBKDF2 해시로 검증.</summary>
    Write = 2,

    /// <summary>열기 + 쓰기 보호 모두 적용.</summary>
    Both  = 3,
}

/// <summary>
/// security/write-lock.json 직렬화 모델.
/// 비밀번호 평문은 저장하지 않고 PBKDF2 파생 해시만 보관한다.
/// 저장 시 사용자 입력을 같은 salt 로 유도해 해시와 상수 시간 비교.
/// </summary>
public sealed class IwpfWriteLock
{
    public int    Version    { get; set; } = 1;
    public string Kdf        { get; set; } = IwpfEncryption.KdfName;
    public int    Iterations { get; set; } = IwpfEncryption.PbkdfIterations;

    /// <summary>PBKDF2 salt (base64).</summary>
    public string Salt { get; set; } = "";

    /// <summary>PBKDF2 출력 해시 (base64, 32바이트).</summary>
    public string Hash { get; set; } = "";
}
