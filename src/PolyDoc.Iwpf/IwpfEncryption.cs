using System.Security.Cryptography;
using System.Text;

namespace PolyDoc.Iwpf;

/// <summary>
/// IWPF 암호화 페이로드. AES-256-GCM 으로 inner ZIP 전체를 암호화한 envelope 형식.
/// 키 유도는 PBKDF2-HMAC-SHA256, 반복 횟수 200,000회.
///
/// 디스크 형식 (security/envelope.json):
/// <code>
/// {
///   "version": 1,
///   "algorithm": "AES-256-GCM",
///   "kdf": "PBKDF2-HMAC-SHA256",
///   "iterations": 200000,
///   "salt": "<base64>",
///   "nonce": "<base64>",
///   "tag": "<base64>"
/// }
/// </code>
/// 암호문 자체는 security/payload.bin 에 raw bytes 로 저장된다.
/// </summary>
public static class IwpfEncryption
{
    public const string AlgorithmName    = "AES-256-GCM";
    public const string KdfName          = "PBKDF2-HMAC-SHA256";
    public const int    PbkdfIterations  = 200_000;
    public const int    KeySize          = 32; // AES-256
    public const int    SaltSize         = 16;
    public const int    NonceSize        = 12; // AES-GCM standard
    public const int    TagSize          = 16;

    /// <summary>지정한 비밀번호로 plaintext 를 암호화. 새 salt/nonce 를 매번 생성.</summary>
    public static (byte[] cipherText, IwpfEnvelope envelope) Encrypt(byte[] plaintext, string password)
    {
        ArgumentNullException.ThrowIfNull(plaintext);
        ArgumentException.ThrowIfNullOrEmpty(password);

        var salt  = RandomNumberGenerator.GetBytes(SaltSize);
        var nonce = RandomNumberGenerator.GetBytes(NonceSize);
        var key   = DeriveKey(password, salt);

        var cipher = new byte[plaintext.Length];
        var tag    = new byte[TagSize];

        using (var aes = new AesGcm(key, TagSize))
        {
            aes.Encrypt(nonce, plaintext, cipher, tag);
        }
        CryptographicOperations.ZeroMemory(key);

        var envelope = new IwpfEnvelope
        {
            Salt  = Convert.ToBase64String(salt),
            Nonce = Convert.ToBase64String(nonce),
            Tag   = Convert.ToBase64String(tag),
        };
        return (cipher, envelope);
    }

    /// <summary>envelope + ciphertext 를 복호화. 잘못된 비밀번호 또는 변조 시 throws.</summary>
    /// <exception cref="WrongIwpfPasswordException">비밀번호가 틀렸거나 페이로드가 변조됨.</exception>
    public static byte[] Decrypt(byte[] cipherText, IwpfEnvelope envelope, string password)
    {
        ArgumentNullException.ThrowIfNull(cipherText);
        ArgumentNullException.ThrowIfNull(envelope);
        ArgumentException.ThrowIfNullOrEmpty(password);

        if (envelope.Algorithm != AlgorithmName)
            throw new InvalidDataException($"Unsupported encryption algorithm '{envelope.Algorithm}'.");
        if (envelope.Kdf != KdfName)
            throw new InvalidDataException($"Unsupported KDF '{envelope.Kdf}'.");

        var salt  = Convert.FromBase64String(envelope.Salt);
        var nonce = Convert.FromBase64String(envelope.Nonce);
        var tag   = Convert.FromBase64String(envelope.Tag);
        var key   = DeriveKey(password, salt, envelope.Iterations);

        var plain = new byte[cipherText.Length];
        try
        {
            using var aes = new AesGcm(key, TagSize);
            aes.Decrypt(nonce, cipherText, tag, plain);
        }
        catch (AuthenticationTagMismatchException)
        {
            CryptographicOperations.ZeroMemory(key);
            throw new WrongIwpfPasswordException();
        }
        catch (CryptographicException)
        {
            CryptographicOperations.ZeroMemory(key);
            throw new WrongIwpfPasswordException();
        }
        CryptographicOperations.ZeroMemory(key);
        return plain;
    }

    /// <summary>비밀번호로 쓰기 잠금 레코드를 생성. 새 salt 를 매번 생성.</summary>
    public static IwpfWriteLock CreateWriteLock(string password)
    {
        ArgumentException.ThrowIfNullOrEmpty(password);
        var salt = RandomNumberGenerator.GetBytes(SaltSize);
        var hash = DeriveKey(password, salt);
        return new IwpfWriteLock
        {
            Salt = Convert.ToBase64String(salt),
            Hash = Convert.ToBase64String(hash),
        };
    }

    /// <summary>비밀번호가 쓰기 잠금 레코드와 일치하는지 상수 시간으로 비교한다.</summary>
    public static bool VerifyWriteLock(string password, IwpfWriteLock writeLock)
    {
        ArgumentException.ThrowIfNullOrEmpty(password);
        ArgumentNullException.ThrowIfNull(writeLock);

        var salt         = Convert.FromBase64String(writeLock.Salt);
        var expectedHash = Convert.FromBase64String(writeLock.Hash);
        var actualHash   = DeriveKey(password, salt, writeLock.Iterations);
        var match        = CryptographicOperations.FixedTimeEquals(actualHash, expectedHash);
        CryptographicOperations.ZeroMemory(actualHash);
        return match;
    }

    private static byte[] DeriveKey(string password, byte[] salt, int iterations = PbkdfIterations)
    {
        return Rfc2898DeriveBytes.Pbkdf2(
            password: Encoding.UTF8.GetBytes(password),
            salt:     salt,
            iterations: iterations,
            hashAlgorithm: HashAlgorithmName.SHA256,
            outputLength: KeySize);
    }
}

/// <summary>security/envelope.json 직렬화 모델.</summary>
public sealed class IwpfEnvelope
{
    public int    Version    { get; set; } = 1;
    public string Algorithm  { get; set; } = IwpfEncryption.AlgorithmName;
    public string Kdf        { get; set; } = IwpfEncryption.KdfName;
    public int    Iterations { get; set; } = IwpfEncryption.PbkdfIterations;
    public string Salt       { get; set; } = "";
    public string Nonce      { get; set; } = "";
    public string Tag        { get; set; } = "";
}

/// <summary>이 IWPF 패키지는 암호화돼 있어 비밀번호 없이 열 수 없음을 알린다.</summary>
public sealed class EncryptedIwpfException : Exception
{
    public EncryptedIwpfException()
        : base("This IWPF package is encrypted. A password is required to open it.") { }
}

/// <summary>비밀번호가 틀렸거나 페이로드가 변조됐음을 알린다.</summary>
public sealed class WrongIwpfPasswordException : Exception
{
    public WrongIwpfPasswordException()
        : base("The provided password is incorrect, or the encrypted payload has been tampered with.") { }
}
