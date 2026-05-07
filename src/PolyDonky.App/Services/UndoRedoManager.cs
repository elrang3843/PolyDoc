using System;
using System.Collections.Generic;
using System.Text.Json;
using PolyDonky.Core;

namespace PolyDonky.App.Services;

/// <summary>
/// 문서 모델 스냅샷 기반 Undo/Redo 매니저.
/// <para>
/// 각 스냅샷은 <see cref="PolyDonkyument"/> 의 JSON 직렬화 결과(byte[]). 깊은 복사를 직렬화로 대체.
/// 스택 깊이 제한(<see cref="MaxDepth"/>) 을 넘으면 가장 오래된 스냅샷부터 폐기.
/// </para>
/// </summary>
public sealed class UndoRedoManager
{
    /// <summary>스택 최대 깊이. 초과 시 바닥부터 폐기. 메모리 폭주 방지용.</summary>
    public const int MaxDepth = 100;

    private readonly LinkedList<byte[]> _undo = new();
    private readonly LinkedList<byte[]> _redo = new();

    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    /// <summary>스택 상태가 바뀔 때마다 발생. UI 의 메뉴 활성/비활성 갱신용.</summary>
    public event EventHandler? StateChanged;

    /// <summary>현재 문서를 undo 스택에 추가하고 redo 스택을 비운다.</summary>
    public void PushUndo(PolyDonkyument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var bytes = Serialize(document);
        _undo.AddLast(bytes);
        TrimToMaxDepth(_undo);
        _redo.Clear();
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// 직전 스냅샷을 꺼내 반환하고, 현재 문서를 redo 스택에 추가한다.
    /// undo 스택이 비어 있으면 <c>null</c>.
    /// </summary>
    public PolyDonkyument? Undo(PolyDonkyument current)
    {
        ArgumentNullException.ThrowIfNull(current);
        if (_undo.Count == 0) return null;

        var currentBytes = Serialize(current);
        var prevBytes    = _undo.Last!.Value;
        _undo.RemoveLast();
        _redo.AddLast(currentBytes);
        TrimToMaxDepth(_redo);
        StateChanged?.Invoke(this, EventArgs.Empty);
        return Deserialize(prevBytes);
    }

    /// <summary>
    /// redo 스택의 다음 스냅샷을 반환하고, 현재 문서를 undo 스택에 추가한다.
    /// redo 스택이 비어 있으면 <c>null</c>.
    /// </summary>
    public PolyDonkyument? Redo(PolyDonkyument current)
    {
        ArgumentNullException.ThrowIfNull(current);
        if (_redo.Count == 0) return null;

        var currentBytes = Serialize(current);
        var nextBytes    = _redo.Last!.Value;
        _redo.RemoveLast();
        _undo.AddLast(currentBytes);
        TrimToMaxDepth(_undo);
        StateChanged?.Invoke(this, EventArgs.Empty);
        return Deserialize(nextBytes);
    }

    /// <summary>두 스택 모두 비우고 상태 변경 이벤트를 한 번 발생시킨다. 문서 새로 열기/저장 직후 호출.</summary>
    public void Clear()
    {
        bool changed = _undo.Count > 0 || _redo.Count > 0;
        _undo.Clear();
        _redo.Clear();
        if (changed) StateChanged?.Invoke(this, EventArgs.Empty);
    }

    private static byte[] Serialize(PolyDonkyument doc)
        => JsonSerializer.SerializeToUtf8Bytes(doc, JsonDefaults.Options);

    private static PolyDonkyument Deserialize(byte[] bytes)
        => JsonSerializer.Deserialize<PolyDonkyument>(bytes, JsonDefaults.Options)
           ?? throw new InvalidOperationException("Undo 스냅샷 역직렬화 실패");

    private static void TrimToMaxDepth(LinkedList<byte[]> list)
    {
        while (list.Count > MaxDepth) list.RemoveFirst();
    }
}
