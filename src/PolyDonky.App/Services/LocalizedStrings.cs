using System.Collections;
using System.ComponentModel;
using System.Globalization;
using SR = PolyDonky.App.Properties.Resources;

namespace PolyDonky.App.Services;

/// <summary>
/// XAML 바인딩용 다국어 문자열 싱글톤.
/// <c>{Binding MenuFile, Source={x:Static svc:LocalizedStrings.Instance}}</c> 형태로 사용.
///
/// ICustomTypeDescriptor 구현으로 Resources.resx 에 있는 모든 문자열 키를
/// 프로퍼티 선언 없이 자동으로 바인딩한다.
/// 새 키는 .resx 에 추가하기만 하면 되고, 이 파일을 편집할 필요가 없다.
///
/// LanguageService.Apply 호출 후 Refresh() 를 실행하면
/// INotifyPropertyChanged.PropertyChanged(string.Empty) 가 전체 갱신을 알려
/// 모든 바인딩이 새 언어 문자열로 재평가된다.
/// </summary>
public sealed class LocalizedStrings : INotifyPropertyChanged, ICustomTypeDescriptor
{
    public static LocalizedStrings Instance { get; } = new();
    private LocalizedStrings() { }

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>언어 변경 후 호출 — 모든 바인딩 대상에 재평가 요청.</summary>
    public void Refresh()
    {
        _cachedProps = null;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(string.Empty));
    }

    // ── 코드 비하인드용 직접 조회 ─────────────────────────────────

    /// <summary>코드 비하인드에서 직접 키로 조회할 때 사용 (SR.xxx 와 동등).</summary>
    internal static string Get(string key) =>
        SR.ResourceManager.GetString(key, SR.Culture ?? CultureInfo.CurrentUICulture) ?? key;

    // ── ICustomTypeDescriptor ────────────────────────────────────

    private PropertyDescriptorCollection? _cachedProps;

    public PropertyDescriptorCollection GetProperties()
    {
        _cachedProps ??= BuildProperties();
        return _cachedProps;
    }

    public PropertyDescriptorCollection GetProperties(Attribute[]? attributes) => GetProperties();

    public AttributeCollection GetAttributes() => AttributeCollection.Empty;
    public string? GetClassName() => null;
    public string? GetComponentName() => null;
    public TypeConverter? GetConverter() => null;
    public EventDescriptor? GetDefaultEvent() => null;
    public PropertyDescriptor? GetDefaultProperty() => null;
    public object? GetEditor(Type editorBaseType) => null;
    public EventDescriptorCollection GetEvents() => EventDescriptorCollection.Empty;
    public EventDescriptorCollection GetEvents(Attribute[]? attributes) => EventDescriptorCollection.Empty;
    public object? GetPropertyOwner(PropertyDescriptor? pd) => this;

    private static PropertyDescriptorCollection BuildProperties()
    {
        // InvariantCulture 로 기본(한국어) 리소스 셋의 키 목록을 가져온다.
        var set = SR.ResourceManager.GetResourceSet(CultureInfo.InvariantCulture, true, true);
        if (set is null) return PropertyDescriptorCollection.Empty;

        var list = new List<PropertyDescriptor>();
        foreach (DictionaryEntry entry in set)
            if (entry.Key is string key && entry.Value is string)
                list.Add(new ResourcePropertyDescriptor(key));

        return new PropertyDescriptorCollection(list.ToArray());
    }
}

internal sealed class ResourcePropertyDescriptor : PropertyDescriptor
{
    public ResourcePropertyDescriptor(string name) : base(name, null) { }

    public override Type ComponentType  => typeof(LocalizedStrings);
    public override bool IsReadOnly     => true;
    public override Type PropertyType   => typeof(string);

    public override bool   CanResetValue(object component)              => false;
    public override void   ResetValue(object component)                 { }
    public override void   SetValue(object? component, object? value)   { }
    public override bool   ShouldSerializeValue(object component)       => false;

    public override object GetValue(object? component) =>
        SR.ResourceManager.GetString(Name, SR.Culture ?? CultureInfo.CurrentUICulture) ?? Name;
}
